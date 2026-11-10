/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The workbook theme as the window stores it (inside the workbook's view
 * options), shared by the window and the spreadsheet editor commands.
 */

import type { SpreadsheetViewOptions } from "./view-options";
import {
  normalizeSpreadsheetWorkbookTheme,
  type SpreadsheetWorkbookThemeSelection,
} from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";

/** The Cell Styles gallery entries that take the theme's heading font. */
export const THEME_HEADING_CELL_STYLE_IDS = new Set([
  "title",
  "heading1",
  "heading2",
  "heading3",
  "heading4",
]);

export function workbookThemeFromViewOptions(
  options: SpreadsheetViewOptions,
): SpreadsheetWorkbookThemeSelection {
  return normalizeSpreadsheetWorkbookTheme({
    colorId: options.themeId,
    fontId: options.themeFontSetId,
    effectId: options.themeEffectId,
    nativeColors: options.nativeThemeColors,
    nativeFonts: options.nativeThemeFonts,
    nativeEffects: options.nativeThemeEffects,
    ...(options.nativeThemeDisposition === "preserve-only"
      ? {
          preserved: {
            ...(options.nativeThemeAccents
              ? { accents: options.nativeThemeAccents }
              : {}),
            ...(options.nativeThemeReplaceable ? { replaceable: true as const } : {}),
          },
        }
      : {}),
  });
}

export function viewOptionsWithWorkbookTheme(
  options: SpreadsheetViewOptions,
  theme: SpreadsheetWorkbookThemeSelection,
): SpreadsheetViewOptions {
  const normalized = normalizeSpreadsheetWorkbookTheme(theme);
  const next: SpreadsheetViewOptions = {
    ...options,
    themeId: normalized.colorId,
  };
  delete next.themeFontSetId;
  delete next.themeEffectId;
  delete next.nativeThemeColors;
  delete next.nativeThemeFonts;
  delete next.nativeThemeEffects;
  delete next.nativeThemeDisposition;
  delete next.nativeThemeAccents;
  delete next.nativeThemeReplaceable;
  if (normalized.preserved) {
    // Back to the theme the file has (Undo of an applied theme).
    next.themeId = "office";
    next.nativeThemeDisposition = "preserve-only";
    if (normalized.preserved.accents) {
      next.nativeThemeAccents = [...normalized.preserved.accents];
    }
    if (normalized.preserved.replaceable) next.nativeThemeReplaceable = true;
    return next;
  }
  if (normalized.nativeColors) next.nativeThemeColors = normalized.nativeColors;
  if (normalized.nativeFonts) next.nativeThemeFonts = normalized.nativeFonts;
  else if (normalized.fontId !== "office") next.themeFontSetId = normalized.fontId;
  if (normalized.nativeEffects) next.nativeThemeEffects = normalized.nativeEffects;
  else if (normalized.effectId !== "none") next.themeEffectId = normalized.effectId;
  return next;
}
