/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's chart templates (XLS-A081 Save as Template), kept per workbook
 * (`chartTemplates` at the content top level). The model and its rules are
 * the editors' shared ones (`lib/office-charts/chart-templates`).
 */
export {
  applyChartTemplate,
  chartTemplateFromModel,
  MAX_CHART_TEMPLATE_NAME_LENGTH,
  MAX_OFFICE_CHART_TEMPLATES as MAX_SPREADSHEET_CHART_TEMPLATES,
  normalizeChartTemplates,
} from "@mrnote/core/web/lib/office-charts/chart-templates.ts";
export type { OfficeChartTemplate as SpreadsheetChartTemplate } from "@mrnote/core/web/lib/office-charts/chart-templates.ts";
