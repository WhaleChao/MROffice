/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Theme colours on a slide object: the deck's Theme Colors grid, and the
 * theme reference a colour picked from it keeps beside the colour.
 *
 * PowerPoint stores a theme swatch as `a:schemeClr` (+ `a:lumMod`/`a:lumOff`),
 * so the object follows the theme when the theme changes; a Standard or
 * custom colour is a plain `a:srgbClr` and stays put (Microsoft Support,
 * "Change the colors in a text box or shape"). The block keeps storing the
 * hex it draws -- every reader and writer still reads that -- and gains
 * `themeColors`, one entry per colour field the reference stands behind:
 *
 *   { fill: { slot: "accent1", lumMod: 20000, lumOff: 80000, rgb: "#c0e4f5" } }
 *
 * An entry is AUTHENTIC only while `rgb` equals the field's own value: an
 * edit that changed the colour some other way (an AI edit, a paste, an
 * older client) leaves an entry that no reader honours, and the next local
 * commit drops it. A deck without the field is exactly what it was.
 */

import type {
  PresentationBlock,
} from "../components/console/editor/native-document/documentCommands";
import { normalizeHexColor } from "./office-color-palette";
import {
  OFFICE_THEME_GRID_SLOTS,
  normalizeOfficeThemeColorRef,
  officeThemeRefHex,
  type OfficeThemeColorRef,
  type OfficeThemeGridColors,
  type OfficeThemeGridSlot,
} from "./office-theme-color-grid";
import type { DeckTheme, PresentationThemeColorSlot } from "./presentation-theme";
import type { OfficeTextPaint } from "./office-text-paint";
import { presentationTextPaintUnderTheme } from "./presentation-wordart-gallery";
import { officeMasterColorMapFlipped, officeThemeMasterIsDark } from "./presentation-office-background-styles";

export const PRESENTATION_THEME_COLORS_KEY = "themeColors";

/** The colour fields a theme reference may stand behind, per block kind. */
export const PRESENTATION_THEME_COLOR_FIELDS = {
  shape: ["fill", "stroke"],
  text: ["fill", "outlineColor"],
  // An SVG graphic's Graphics Fill / Outline (PPT-A161): PowerPoint writes a
  // theme swatch as the SVG's `MsftOfcThm_*` class and follows the theme.
  officeGraphic: ["graphicFill", "graphicOutline"],
} as const;

export type PresentationThemeColorField =
  | (typeof PRESENTATION_THEME_COLOR_FIELDS)["shape"][number]
  | (typeof PRESENTATION_THEME_COLOR_FIELDS)["text"][number]
  | (typeof PRESENTATION_THEME_COLOR_FIELDS)["officeGraphic"][number];

export interface PresentationThemeColorBinding extends OfficeThemeColorRef {
  /** The colour the reference drew when it was picked: its authentication. */
  readonly rgb: string;
}

export type PresentationThemeColorBindings = Partial<
  Record<PresentationThemeColorField, PresentationThemeColorBinding>
>;

/**
 * The theme colours a style recipe paints with: each slot's own colour, so a
 * Shape Styles pick binds the fill or outline that IS a theme colour (Accent
 * 1, Background 1) as PowerPoint's style references do.
 */
export function presentationThemeSlotPicks(
  colors: OfficeThemeGridColors,
): PresentationThemeColorPick[] {
  const order: OfficeThemeGridSlot[] = [
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "tx1", "bg1", "tx2", "bg2",
  ];
  return order.map((slot) => ({ color: colors[slot], ref: { slot } }));
}

/** What a palette just picked: the colour, and its theme reference if any. */
export interface PresentationThemeColorPick {
  readonly color: string | null;
  readonly ref: OfficeThemeColorRef | null;
}

/** Office 2013-2022's colours, the deck's fallback for a slot it never states
 * (the slots a theme-less deck's download carries from its template). */
const FALLBACK_SLOTS: OfficeThemeGridColors = {
  bg1: "#ffffff",
  tx1: "#000000",
  bg2: "#e7e6e6",
  tx2: "#44546a",
  accent1: "#4472c4",
  accent2: "#ed7d31",
  accent3: "#a5a5a5",
  accent4: "#ffc000",
  accent5: "#5b9bd5",
  accent6: "#70ad47",
};

const GRID_TO_THEME_SLOT: Record<OfficeThemeGridSlot, PresentationThemeColorSlot> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
};

/** The dark colour map, the one Office's dark variants use. */
const DARK_MAP_THEME_SLOT: Record<OfficeThemeGridSlot, PresentationThemeColorSlot> = {
  ...GRID_TO_THEME_SLOT,
  bg1: "dk1",
  tx1: "lt1",
  bg2: "dk2",
  tx2: "lt2",
};

/**
 * The ten picker colours of a deck's theme: its twelve Office slots when it
 * states them (an Office theme, or an imported deck's own theme), else the
 * colours its download writes into those slots (background, title and body
 * colours, accents), else Office's.
 */
export function presentationThemeGridColors(
  theme: DeckTheme | null | undefined,
): OfficeThemeGridColors {
  const identity = theme?.identity;
  const slots =
    theme?.colorSlots ??
    (identity?.selectionKind === "native" && identity.variantId === "base"
      ? identity.source?.colorSlots
      : undefined);
  // bg1 / tx1 / bg2 / tx2 name a slot through the master's colour map (`p:clrMap`):
  // a dark variant's Background 1 is its dark slot. Read as the light map, a
  // reselected dark theme turned its master's lines from #262626 to #ffffff, and
  // the Theme Colors grid offered the other end of each pair.
  const darkMap = officeThemeMasterIsDark(theme) !== officeMasterColorMapFlipped(theme);
  const stated = (slot: OfficeThemeGridSlot): string | null =>
    normalizeHexColor(slots?.[darkMap ? DARK_MAP_THEME_SLOT[slot] : GRID_TO_THEME_SLOT[slot]]);
  const accent = (index: number) =>
    normalizeHexColor(theme?.accentColors?.[index]);
  const colors = {} as Record<OfficeThemeGridSlot, string>;
  for (const slot of OFFICE_THEME_GRID_SLOTS) {
    const fromFields =
      slot === "bg1"
        ? normalizeHexColor(theme?.background)
        : slot === "tx1"
          ? normalizeHexColor(theme?.titleColor)
          : slot === "tx2"
            ? normalizeHexColor(theme?.bodyTextColor)
            : slot.startsWith("accent")
              ? accent(Number(slot.slice(6)) - 1)
              : null;
    colors[slot] = stated(slot) ?? fromFields ?? FALLBACK_SLOTS[slot];
  }
  return colors;
}

export function normalizePresentationThemeColorBindings(
  value: unknown,
): PresentationThemeColorBindings | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const result: Record<string, PresentationThemeColorBinding> = {};
  for (const [field, raw] of Object.entries(value as Record<string, unknown>)) {
    if (!["fill", "stroke", "outlineColor", "graphicFill", "graphicOutline"].includes(field)) continue;
    const ref = normalizeOfficeThemeColorRef(raw);
    const rgb = normalizeHexColor((raw as { rgb?: unknown } | null)?.rgb);
    if (ref && rgb) result[field] = { ...ref, rgb };
  }
  return Object.keys(result).length ? result : undefined;
}

function blockFields(block: PresentationBlock): readonly PresentationThemeColorField[] {
  if (block.type === "shape") return PRESENTATION_THEME_COLOR_FIELDS.shape;
  if (block.type === "text") return PRESENTATION_THEME_COLOR_FIELDS.text;
  if (block.type === "image" && block.officeGraphicKind === "svg") {
    return PRESENTATION_THEME_COLOR_FIELDS.officeGraphic;
  }
  return [];
}

/** The binding behind one field, only while it still authenticates. */
export function presentationThemeColorBinding(
  block: PresentationBlock,
  field: PresentationThemeColorField,
): PresentationThemeColorBinding | null {
  const binding = normalizePresentationThemeColorBindings(
    (block as Record<string, unknown>)[PRESENTATION_THEME_COLORS_KEY],
  )?.[field];
  const value = normalizeHexColor((block as Record<string, unknown>)[field]);
  return binding && value === binding.rgb ? binding : null;
}

function withBindings(
  block: PresentationBlock,
  bindings: PresentationThemeColorBindings,
): PresentationBlock {
  const next = { ...block } as Record<string, unknown>;
  if (Object.keys(bindings).length) next[PRESENTATION_THEME_COLORS_KEY] = bindings;
  else delete next[PRESENTATION_THEME_COLORS_KEY];
  return next as PresentationBlock;
}

/**
 * One edit's effect on a block's references. A field whose colour the edit
 * changed to exactly what a theme swatch just picked takes that swatch's
 * reference; a field that changed any other way loses its reference; a
 * reference the writer set itself (its `rgb` matches) stays.
 */
export function reconcilePresentationThemeColors(
  before: PresentationBlock | undefined,
  after: PresentationBlock,
  pick: PresentationThemeColorPick | readonly PresentationThemeColorPick[] | null,
): PresentationBlock {
  const fields = blockFields(after);
  if (!fields.length) return after;
  const record = after as Record<string, unknown>;
  const previous = before as Record<string, unknown> | undefined;
  const stored = normalizePresentationThemeColorBindings(
    record[PRESENTATION_THEME_COLORS_KEY],
  );
  const bindings: PresentationThemeColorBindings = {};
  const picks = (Array.isArray(pick) ? pick : pick ? [pick] : []) as readonly PresentationThemeColorPick[];
  const pickFor = (value: string | null) =>
    value
      ? picks.find((candidate) => candidate.ref && normalizeHexColor(candidate.color) === value)
      : undefined;
  for (const field of fields) {
    const value = normalizeHexColor(record[field]);
    const own = stored?.[field];
    const changed = !previous || previous[field] !== record[field];
    const picked = changed ? pickFor(value) : undefined;
    if (own && value === own.rgb && !picked) {
      bindings[field] = own;
    } else if (picked?.ref && value) {
      bindings[field] = { ...picked.ref, rgb: value };
    }
  }
  if (
    JSON.stringify(stored ?? {}) === JSON.stringify(bindings) &&
    (stored !== undefined || !(PRESENTATION_THEME_COLORS_KEY in record))
  ) {
    return after;
  }
  return withBindings(after, bindings);
}

/**
 * The selected blocks with a theme swatch's reference bound to every colour
 * field that already shows the swatch's RGB: a swatch picked over a plain
 * colour of the same RGB writes no field, so nothing else records that the
 * colour is the theme's from here on (PPT-B427). The blocks rewritten, or
 * null when no selected block carries a picked colour.
 */
export function presentationBlocksWithPickedThemeColors(
  blocks: readonly PresentationBlock[],
  ids: ReadonlySet<string>,
  pick: PresentationThemeColorPick | readonly PresentationThemeColorPick[],
): PresentationBlock[] | null {
  let changed = false;
  const next = blocks.map((block) => {
    if (!ids.has(block.id)) return block;
    const bound = reconcilePresentationThemeColors(undefined, block, pick);
    if (bound !== block) changed = true;
    return bound;
  });
  return changed ? next : null;
}

/** {@link reconcilePresentationThemeColors} over a slide's blocks, by id. */
export function reconcilePresentationBlocksThemeColors(
  before: readonly PresentationBlock[],
  after: PresentationBlock[],
  pick: PresentationThemeColorPick | readonly PresentationThemeColorPick[] | null,
): PresentationBlock[] {
  const previous = new Map(before.map((block) => [block.id, block]));
  let changed = false;
  const next = after.map((block) => {
    const old = previous.get(block.id);
    if (old === block) return block;
    const reconciled = reconcilePresentationThemeColors(old, block, pick);
    if (reconciled !== block) changed = true;
    return reconciled;
  });
  return changed ? next : after;
}

/**
 * The deck under a NEW theme: every authentic reference is resolved again,
 * so a shape filled with "Accent 1, Lighter 80%" takes the new theme's
 * Accent 1, Lighter 80% -- as PowerPoint recolours theme-coloured objects
 * when the theme changes. Plain colours and stale references stay as they are.
 *
 * `previousColors`, for a deck whose colour map is not known from an Office theme
 * (an imported deck's own master): a reference whose stored colour is what it
 * resolves to under the opposite map was resolved under that map, so it goes on
 * being resolved so.
 */
export function presentationBlockUnderTheme(
  input: PresentationBlock,
  colors: OfficeThemeGridColors,
  previousColors?: OfficeThemeGridColors,
): PresentationBlock {
  const block = wordArtUnderTheme(input, colors);
  const fields = blockFields(block);
  if (!fields.length) return block;
  const bindings: PresentationThemeColorBindings = {};
  const patch: Record<string, string> = {};
  let changed = false;
  for (const field of fields) {
    const binding = presentationThemeColorBinding(block, field);
    if (!binding) continue;
    const opposite = previousColors && resolvedUnderOppositeMap(binding, previousColors);
    const resolved = officeThemeRefHex(binding, opposite ? oppositeMapColors(colors) : colors);
    if (!resolved) {
      bindings[field] = binding;
      continue;
    }
    bindings[field] = { ...binding, rgb: resolved };
    if (resolved !== binding.rgb) {
      patch[field] = resolved;
      changed = true;
    }
  }
  if (!changed) return block;
  return withBindings({ ...block, ...patch } as PresentationBlock, bindings);
}

function oppositeMapColors(colors: OfficeThemeGridColors): OfficeThemeGridColors {
  return { ...colors, bg1: colors.tx1, tx1: colors.bg1, bg2: colors.tx2, tx2: colors.bg2 };
}

/** Transforms round on the file's side and on ours: a step either way is the same colour. */
function sameColorWithinRounding(a: string | null, b: string | null): boolean {
  if (!a || !b) return false;
  const parts = (hex: string) => [1, 3, 5].map((at) => Number.parseInt(hex.slice(at, at + 2), 16));
  const [left, right] = [parts(a), parts(b)];
  return left.every((value, index) => Math.abs(value - right[index]!) <= 2);
}

function resolvedUnderOppositeMap(
  binding: PresentationThemeColorBinding,
  previous: OfficeThemeGridColors,
): boolean {
  if (!["bg1", "tx1", "bg2", "tx2"].includes(binding.slot)) return false;
  return !sameColorWithinRounding(officeThemeRefHex(binding, previous), binding.rgb) &&
    sameColorWithinRounding(officeThemeRefHex(binding, oppositeMapColors(previous)), binding.rgb);
}

/** A text box's or shape's WordArt under the new theme (PPT-A059). */
function wordArtUnderTheme(block: PresentationBlock, colors: OfficeThemeGridColors): PresentationBlock {
  if (block.type !== "text" && block.type !== "shape") return block;
  const paint = (block as { textPaint?: OfficeTextPaint }).textPaint;
  if (!paint?.stated) return block;
  const next = presentationTextPaintUnderTheme(paint, colors);
  return JSON.stringify(next) === JSON.stringify(paint)
    ? block
    : ({ ...block, textPaint: next } as PresentationBlock);
}
