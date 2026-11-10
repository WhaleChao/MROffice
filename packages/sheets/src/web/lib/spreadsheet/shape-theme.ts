/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** DrawingML style-matrix projection. The source XML is never mutated. */
import {
  officeColorChannelsFromHex,
  officeColorHexFromChannels,
  officeColorTransformChannels,
  type OfficeColorChannels,
} from "@mrnote/core/web/lib/office-color-transform.ts";
import {
  officeShapeStyleRecipes,
  type OfficeShapeStyleRecipe,
} from "@mrnote/core/web/lib/office-shape-style-catalog.ts";
type Transform = readonly [string, number];
const ACCEPTED = new Set(["tint", "shade", "lumMod", "lumOff", "satMod", "satOff"]);

/**
 * One colour after its ordered transforms. The members accepted stay the ones
 * the server's twin accepts (`drawing_color`); the arithmetic is the rule every
 * host asks (`office-color-transform`).
 */
export function spreadsheetDrawingColor(
  hex: string,
  transforms: readonly Transform[],
): string {
  const base = /^#[0-9a-f]{6}$/i.test(hex)
    ? officeColorChannelsFromHex(hex)
    : null;
  if (!base) throw Error("unsupported_shape_theme_color");
  let channels: OfficeColorChannels = base;
  for (const [name, raw] of transforms) {
    if (!Number.isFinite(raw) || raw < 0 || raw > 1000000)
      throw Error("unsupported_shape_color_transform");
    if ((name === "shade" || name === "tint") && raw > 100000)
      throw Error("unsupported_shape_color_transform");
    const moved: OfficeColorChannels | null = ACCEPTED.has(name)
      ? officeColorTransformChannels(channels, name, raw)
      : null;
    if (!moved) throw Error("unsupported_shape_color_transform");
    channels = moved;
  }
  return officeColorHexFromChannels(channels);
}

const subtle: readonly Transform[][] = [
  [
    ["lumMod", 110000],
    ["satMod", 105000],
    ["tint", 67000],
  ],
  [
    ["lumMod", 105000],
    ["satMod", 103000],
    ["tint", 73000],
  ],
  [
    ["lumMod", 105000],
    ["satMod", 109000],
    ["tint", 81000],
  ],
];
const moderate: readonly Transform[][] = [
  [
    ["satMod", 103000],
    ["lumMod", 102000],
    ["tint", 94000],
  ],
  [
    ["satMod", 110000],
    ["lumMod", 100000],
    ["shade", 100000],
  ],
  [
    ["lumMod", 99000],
    ["satMod", 120000],
    ["shade", 78000],
  ],
];

/** Same shared gallery/IDs, with native Office paint units, not thumbnail pixels. */
export function spreadsheetNativeShapeStyleRecipes(
  accents: readonly string[],
): readonly OfficeShapeStyleRecipe[] {
  return officeShapeStyleRecipes(accents).map((recipe) => {
    const color =
      recipe.column === 0
        ? "#000000"
        : (accents[recipe.column - 1] ?? recipe.stroke);
    if (recipe.row === "outline") return { ...recipe, stroke: color };
    if (recipe.row === "outlined-fill")
      return {
        ...recipe,
        fill: color,
        stroke: spreadsheetDrawingColor(color, [["shade", 15000]]),
      };
    if (recipe.row === "solid") return { ...recipe, fill: color };
    const stops = (recipe.row === "light-gradient" ? subtle : moderate).map(
      (transforms, i) => ({
        position: i / 2,
        color: spreadsheetDrawingColor(color, transforms),
      }),
    );
    return {
      ...recipe,
      fill: stops[0].color,
      fillTo: stops[2].color,
      fillGradient: { angle: 180, stops },
      ...(recipe.row === "shadow"
        ? {
            shapeShadow: {
              color: "#000000",
              blur: 6,
              distance: 2,
              opacity: 0.63,
            },
          }
        : { stroke: color }),
    };
  });
}
