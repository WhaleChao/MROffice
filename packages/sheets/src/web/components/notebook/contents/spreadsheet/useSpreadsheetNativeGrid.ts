"use client";
/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useMemo } from "react";
import { normalizeNativeGrid } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";
import { resolveSpreadsheetNativeGrid } from "../../../../lib/spreadsheet/row-column-buttons";
import { nativeGridFontForLocale, type SpreadsheetThemeFontSet } from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";


















export function useSpreadsheetNativeGrid(raw: unknown, widths: Record<string, number> | undefined, defaultWidth: number | undefined, themeId?: string, bodyFamily?: string, headingFamily?: string, localeFonts?: SpreadsheetThemeFontSet, locale?: string) {
  const grid = useMemo(() => normalizeNativeGrid(raw), [raw]);
  // The grid font as the UI's language draws it: Chinese Excel counts a
  // theme-bound Normal font's columns in the theme's Hans face (X-329).
  const measured = grid?.metricsFont ?? grid?.font;
  const localeFamily = measured && localeFonts && locale ? nativeGridFontForLocale(measured, localeFonts, locale).family : undefined;
  return useMemo(
    () => resolveSpreadsheetNativeGrid(grid, widths, defaultWidth, themeId, bodyFamily, headingFamily, localeFamily),
    [grid, widths, defaultWidth, themeId, bodyFamily, headingFamily, localeFamily],
  );
}
