/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  normalizeOfficeShapeGradient,
  resolveOfficeStopColor,
  type OfficeColorSource,
  type OfficeColorTransformEntry,
  type OfficeShapeGradient,
} from "./office-shape-paint";
import type { OfficeShapeStyleRecipe } from "./office-shape-style-catalog";






























export const OFFICE_SHAPE_STYLE_PRESET_ROWS = [
  "preset-transparent",
  "preset-outline",
  "preset-semitransparent",
  "preset-fill",
  "preset-gradient",
] as const;
export type OfficeShapeStylePresetRow =
  (typeof OFFICE_SHAPE_STYLE_PRESET_ROWS)[number];

export function isOfficeShapeStylePresetRow(
  row: string,
): row is OfficeShapeStylePresetRow {
  return (OFFICE_SHAPE_STYLE_PRESET_ROWS as readonly string[]).includes(row);
}

/** The theme slot each gallery column paints with. */
export const OFFICE_SHAPE_STYLE_PRESET_SLOTS = [
  "dk1",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
] as const;

/** PowerPoint's preset gradient: the stop positions and their transforms. */
const PRESET_GRADIENT_STOPS: readonly (readonly [
  number,
  readonly OfficeColorTransformEntry[],
])[] = [
  [0, [["lumMod", 67_000]]],
  [0.48, [["lumMod", 97_000], ["lumOff", 3_000]]],
  [1, [["lumMod", 60_000], ["lumOff", 40_000]]],
];

function presetGradient(slot: string, base: string): OfficeShapeGradient {
  return normalizeOfficeShapeGradient({
    // `a:lin ang="16200000"` (270 deg) is CSS 0 deg: the first stop at the bottom.
    angle: 0,
    stops: PRESET_GRADIENT_STOPS.map(([position, transforms]) => {
      const source: OfficeColorSource = {
        kind: "scheme",
        val: slot,
        base,
        transforms: transforms.map(
          ([kind, raw]): OfficeColorTransformEntry => [kind, raw],
        ),
      };
      return { position, color: resolveOfficeStopColor(source) ?? base, source };
    }),
    scaled: true,
    rotWithShape: true,
    flip: "none",
    tileRect: {},
  })!;
}

/**
 * The measured presets for one deck: `columns` is the seven colours the
 * gallery's columns draw with (Dark 1, then Accent 1-6).
 */
export function officeShapeStylePresetRecipes(
  columns: readonly string[],
  light1 = "#ffffff",
): readonly OfficeShapeStyleRecipe[] {
  return OFFICE_SHAPE_STYLE_PRESET_ROWS.flatMap((row) =>
    columns.slice(0, 7).map((color, column): OfficeShapeStyleRecipe => {
      const base = {
        id: `${row}-${column}`,
        row,
        column,
        stroke: "none",
        strokeWidth: 0,
        shadow: false,
      };
      if (row === "preset-transparent")
        return { ...base, fill: "none", textColor: color };
      // 9525 EMU, PowerPoint's 0.75 pt line: one CSS pixel in this model.
      if (row === "preset-outline")
        return { ...base, fill: "none", stroke: color, strokeWidth: 1, textColor: color };
      if (row === "preset-semitransparent")
        return { ...base, fill: color, fillTransparency: 50, textColor: light1 };
      if (row === "preset-fill")
        return { ...base, fill: color, textColor: light1 };
      const fillGradient = presetGradient(
        OFFICE_SHAPE_STYLE_PRESET_SLOTS[column],
        color,
      );
      const first = fillGradient.stops[0].color;
      const last = fillGradient.stops[fillGradient.stops.length - 1].color;
      return {
        ...base,
        fill: first,
        fillTo: last,
        gradientAngle: 0,
        fillGradient,
        textColor: light1,
      };
    }),
  );
}
