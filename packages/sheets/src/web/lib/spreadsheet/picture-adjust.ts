/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A sheet picture's Picture Format > Color, Corrections and Transparency as
 * the grid draws them: the shared picture filter (`office-picture-adjust-filter`)
 * in the workbook's theme, and the picture's opacity. The object keeps what
 * its blip states (`pictureRecolor`, `pictureCorrections`,
 * `pictureTransparency`); a picture that states none draws as it is.
 */

import type { OfficeThemeGridColors } from "@mrnote/core/web/lib/office-theme-color-grid.ts";
import { officePictureAdjustFilterCss } from "@mrnote/core/web/lib/office-picture-adjust-filter.ts";
import { isOfficeRecolorPreset } from "@mrnote/core/web/lib/office-recolor-presets.ts";
import { paletteById } from "@mrnote/core/web/lib/spreadsheet/color-palette.ts";
import {
  normalizeSpreadsheetWorkbookTheme,
  spreadsheetWorkbookThemePaletteId,
} from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";
import type { SpreadsheetWorkbookThemeSelection } from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";

interface PictureAdjustObject {
  readonly pictureRecolor?: unknown;
  readonly pictureCorrections?: unknown;
  readonly pictureTransparency?: unknown;
}

/** The workbook's theme as the recolour grid reads it. */
export function spreadsheetPictureRecolorTheme(
  selection: SpreadsheetWorkbookThemeSelection | null | undefined,
): OfficeThemeGridColors | null {
  if (!selection) return null;
  const normalized = normalizeSpreadsheetWorkbookTheme(selection);
  const colors = normalized.nativeColors?.colors;
  const palette = colors ? null : paletteById(spreadsheetWorkbookThemePaletteId(normalized));
  const slot = (native: keyof NonNullable<typeof colors>, fallback: string | undefined) =>
    (colors ? colors[native] : fallback) ?? "#000000";
  return {
    tx1: slot("dk1", palette?.dark1),
    bg1: slot("lt1", palette?.light1),
    tx2: slot("dk2", palette?.dark2),
    bg2: slot("lt2", palette?.light2),
    accent1: slot("accent1", palette?.accents[0]),
    accent2: slot("accent2", palette?.accents[1]),
    accent3: slot("accent3", palette?.accents[2]),
    accent4: slot("accent4", palette?.accents[3]),
    accent5: slot("accent5", palette?.accents[4]),
    accent6: slot("accent6", palette?.accents[5]),
  };
}

function percent(value: unknown, low: number, high: number): number {
  return typeof value === "number" && Number.isInteger(value) && value >= low && value <= high ? value : 0;
}

/** The CSS filter terms for the picture's Color and Corrections ("" for none). */
export function spreadsheetPictureAdjustFilter(
  object: PictureAdjustObject,
  theme: OfficeThemeGridColors | null,
): string {
  const corrections = (object.pictureCorrections ?? {}) as { brightness?: unknown; contrast?: unknown };
  return officePictureAdjustFilterCss(
    {
      recolor: object.pictureRecolor,
      brightness: percent(corrections.brightness, -100, 100),
      contrast: percent(corrections.contrast, -100, 100),
    },
    theme,
  );
}

/** The picture's opacity under its Transparency (undefined when opaque). */
export function spreadsheetPictureOpacity(object: PictureAdjustObject): number | undefined {
  const transparency = percent(object.pictureTransparency, 1, 100);
  return transparency ? (100 - transparency) / 100 : undefined;
}

/**
 * The object change a Picture Format > Adjust pick makes, or null when it
 * changes nothing: Corrections as `"brightness,contrast"`, Color as a Recolor
 * preset id (or "none"), Transparency as a whole percentage. Washout is a
 * correction of its own: choosing it clears Corrections, and Corrections
 * clear it (the blip has one `a:lum`).
 */
export function spreadsheetPictureAdjustPatch(
  object: PictureAdjustObject,
  kind: "corrections" | "color" | "transparency",
  parameter: string | number,
): {
  pictureRecolor?: string | undefined;
  pictureCorrections?: { brightness: number; contrast: number } | undefined;
  pictureTransparency?: number | undefined;
} | null {
  const recolor = isOfficeRecolorPresetValue(object.pictureRecolor) ? object.pictureRecolor : undefined;
  const current = (object.pictureCorrections ?? null) as { brightness?: unknown; contrast?: unknown } | null;
  const currentPair = current
    ? { brightness: percent(current.brightness, -100, 100), contrast: percent(current.contrast, -100, 100) }
    : { brightness: 0, contrast: 0 };
  if (kind === "corrections") {
    const [brightness, contrast] = String(parameter).split(",").map((value) => Number(value));
    if (!Number.isInteger(brightness) || !Number.isInteger(contrast)) return null;
    if (Math.abs(brightness!) > 100 || Math.abs(contrast!) > 100) return null;
    if (brightness === currentPair.brightness && contrast === currentPair.contrast && recolor !== "washout") return null;
    return {
      pictureCorrections: brightness || contrast ? { brightness: brightness!, contrast: contrast! } : undefined,
      ...(recolor === "washout" ? { pictureRecolor: undefined } : {}),
    };
  }
  if (kind === "color") {
    const id = String(parameter);
    const next = id === "none" ? undefined : isOfficeRecolorPresetValue(id) ? id : null;
    if (next === null || next === recolor) return null;
    return {
      pictureRecolor: next,
      ...(next === "washout" ? { pictureCorrections: undefined } : {}),
    };
  }
  const value = Number(parameter);
  if (!Number.isInteger(value) || value < 0 || value > 100) return null;
  const next = value || undefined;
  if (next === (percent(object.pictureTransparency, 1, 100) || undefined)) return null;
  return { pictureTransparency: next };
}

function isOfficeRecolorPresetValue(value: unknown): value is string {
  return isOfficeRecolorPreset(value);
}
