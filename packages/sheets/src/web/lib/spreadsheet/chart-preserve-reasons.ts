/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Why an imported chart is kept exactly as saved (preserve-only), in words.
 *
 * The import receipt lists the reasons the shared chart reader could not
 * state the chart (`spreadsheetChartPreserveReasons`): the first is the kind
 * of gap (`view-unread`: the reader left something out; `part-graph`,
 * `chart-ex-part`: the part itself), and after `view-unread` come what it left
 * out. The most specific one picks the message every chart command shows,
 * instead of a code.
 */

import { spreadsheetChartPreserveReasons, type SpreadsheetChartLike } from "@mrnote/core/web/lib/spreadsheet/chart-v2-model.ts";

const REASON_KEYS: ReadonlyArray<readonly [RegExp, string]> = [
  [/^title-paint$/, "spreadsheet.uiCopy.chartKeptTitleFormatting"],
  [/^(axis-paint|axis-titles|gridlines)$/, "spreadsheet.uiCopy.chartKeptAxisFormatting"],
  [/^series-paint$/, "spreadsheet.uiCopy.chartKeptSeriesFormatting"],
  [/^model:dataLabel/, "spreadsheet.uiCopy.chartKeptDataLabelOptions"],
  [/^category-levels$/, "spreadsheet.uiCopy.chartKeptCategoryLevels"],
  [/^data-table$/, "spreadsheet.uiCopy.chartKeptDataTable"],
  [/^(error-bars|trendlines)$/, "spreadsheet.uiCopy.chartKeptAnalysis"],
  [
    /^(series-name|text-beyond-model|points-beyond-model|series-beyond-model|point-gaps|pie-series|categories|view-differs)$/,
    "spreadsheet.uiCopy.chartKeptUnreadData",
  ],
  [/^chart-ex-part$/, "spreadsheet.uiCopy.chartKeptOffice2016Chart"],
];

const STRUCTURE_KEY = "spreadsheet.uiCopy.chartKeptStructure";

/** The message key for a preserve-only chart, or null for one the editor
 * may change. */
export function spreadsheetChartPreserveReasonKey(chart: SpreadsheetChartLike): string | null {
  const reasons = spreadsheetChartPreserveReasons(chart);
  if (!reasons) return null;
  const specific = reasons[0] === "view-unread" ? reasons.slice(1) : reasons;
  for (const reason of specific) {
    for (const [pattern, key] of REASON_KEYS) {
      if (pattern.test(reason)) return key;
    }
  }
  return STRUCTURE_KEY;
}
