/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck's chart type menus as PowerPoint for Mac draws them (PPT-A062):
 * Insert > Chart and Chart Design > Change Chart Type list the same nine
 * entries (`OFFICE_MAC_CHART_MENU_GROUPS`), three of them under PowerPoint's
 * own shorter names (Hierarchy, Statistical, X Y (Scatter)).
 */
import type { OfficeChartTypeGroup } from "@mrnote/core/web/components/office-ribbon/OfficeChartTypeMenu.tsx";
import {
  chartFamilyIcon,
  type ChartFamily,
} from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { OFFICE_MAC_CHART_MENU_GROUPS } from "@mrnote/core/web/lib/office-charts/catalog.ts";

const MENU_WORDS: Partial<Record<ChartFamily, string>> = {
  hierarchy: "presentation.insert.chartFamily.hierarchy",
  statistical: "presentation.insert.chartFamily.statistical",
  xy: "presentation.insert.chartFamily.xy",
};

/** A chart family's name on the deck's chart menus. */
export function presentationChartFamilyLabel(
  translate: (key: string) => string,
  family: ChartFamily,
): string {
  return translate(MENU_WORDS[family] ?? `toolbar.chartV2.family.${family}`);
}

/** Change Chart Type's entries: PowerPoint for Mac's nine. */
export function presentationChartTypeGroups(
  translate: (key: string) => string,
): OfficeChartTypeGroup[] {
  return OFFICE_MAC_CHART_MENU_GROUPS.map((group) => ({
    id: group.family,
    label: presentationChartFamilyLabel(translate, group.family),
    icon: chartFamilyIcon(group.family),
    families: group.families,
  }));
}
