/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The worksheet's names for the one Quick Layout catalogue
 * (`lib/office-charts/chart-layout.ts`), and the reader of a column / bar
 * part's layout atoms.
 */
import {
  OFFICE_CHART_QUICK_LAYOUT_IDS,
  normalizeOfficeChartLayout,
  officeChartQuickLayoutRecipe,
  type OfficeChartLabelOverride,
  type OfficeChartLayout,
  type OfficeChartQuickLayoutId,
} from "../office-charts/chart-layout";

export const CHART_QUICK_LAYOUT_IDS = OFFICE_CHART_QUICK_LAYOUT_IDS;
export type ChartQuickLayoutId = OfficeChartQuickLayoutId;
export type SpreadsheetChartLabelOverride = OfficeChartLabelOverride;
export type SpreadsheetChartLayout = OfficeChartLayout;
export const normalizeSpreadsheetChartLayout = normalizeOfficeChartLayout;
export const chartQuickLayoutRecipe = officeChartQuickLayoutRecipe;
