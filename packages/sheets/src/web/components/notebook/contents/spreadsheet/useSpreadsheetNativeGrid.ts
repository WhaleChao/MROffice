"use client";
/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useMemo } from "react";
import { normalizeNativeGrid } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";
import { resolveSpreadsheetNativeGrid } from "../../../../lib/spreadsheet/row-column-buttons";
import { nativeGridFontForLocale, nativeGridLatinLocale, type SpreadsheetThemeFontSet } from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";


















export function useSpreadsheetNativeGrid(raw: unknown, widths: Record<string, number> | undefined, defaultWidth: number | undefined, themeId?: string, bodyFamily?: string, headingFamily?: string, localeFonts?: SpreadsheetThemeFontSet, locale?: string) {
  const grid = useMemo(() => normalizeNativeGrid(raw), [raw]);
  // The grid font as the UI's language draws it: Chinese Excel counts a
  // theme-bound Normal font's columns in the theme's Hans face (X-329).
  const measured = grid?.metricsFont ?? grid?.font;
  const localeFamily = measured && localeFonts && locale ? nativeGridFontForLocale(measured, localeFonts, locale).family : undefined;
  // A Western UI measures rows in that face too (Windows Excel en-US stands a
  // 等线 + minor sheet's rows at Calibri's, lookwin p28); Chinese Excel keeps
  // the workbook's rows (X-329).
  const localeRows = Boolean(locale && nativeGridLatinLocale(locale));
  return useMemo(
    () => resolveSpreadsheetNativeGrid(grid, widths, defaultWidth, themeId, bodyFamily, headingFamily, localeFamily, localeRows),
    [grid, widths, defaultWidth, themeId, bodyFamily, headingFamily, localeFamily, localeRows],
  );
}
