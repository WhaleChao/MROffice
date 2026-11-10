"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useState } from "react";
import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { DropdownMenu, DropdownMenuTrigger } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { OfficeChartTypeMenuItems, type OfficeChartTypeGroup } from "@mrnote/core/web/components/office-ribbon/OfficeChartTypeMenu.tsx";
import { type ChartFamily, type ChartSubtype } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";

export interface SpreadsheetDesktopChartTypeChoice {
  type: string;
  subtype: ChartSubtype;
  enabled: boolean;
  reason?: string;
}
// Excel's own Insert-tab grouping, which the worksheet has always shown; the
// list itself is the one Change Chart Type menu every editor mounts.
const GROUPS = [
  {id:"column",labelKey:"spreadsheet.helperCopy.chartType.column",icon:"chart.type.column",families:["column","bar"]},
  {id:"line",labelKey:"spreadsheet.helperCopy.chartType.line",icon:"chart.type.line",families:["line","area","radar"]},
  {id:"pie",labelKey:"spreadsheet.helperCopy.chartType.pie",icon:"chart.type.pie",families:["pie"]},
  {id:"hierarchy",labelKey:"spreadsheet.helperCopy.chartType.hierarchy",icon:"chart.type.other",families:["hierarchy"]},
  {id:"statistical",labelKey:"spreadsheet.helperCopy.chartType.statistical",icon:"chart.type.statistical",families:["statistical","stock"]},
  {id:"xy",labelKey:"spreadsheet.helperCopy.chartType.xy",icon:"chart.type.scatter",families:["xy"]},
  {id:"waterfall",labelKey:"spreadsheet.helperCopy.chartType.waterfall",icon:"chart.type.other",families:["waterfall","funnel"]},
  {id:"combo",labelKey:"spreadsheet.helperCopy.chartType.combo",icon:"chart.type.combo",families:["combo"]},
  {id:"map",labelKey:"spreadsheet.helperCopy.chartType.map",icon:"chart.type.other",families:["map"]},
] as const;

/** The same command catalogue feeds every row; a native capability owns the
 * selection, not a menu label or a preview-only renderer. */
export function SpreadsheetChartTypeMenu({translate,disabledReason,current,choice,label,onChoose}: {
  translate:(key:string)=>string;
  disabledReason:string|null;
  current:{type:string;subtype?:string};
  choice:(family:ChartFamily,subtype:ChartSubtype)=>SpreadsheetDesktopChartTypeChoice;
  label:(family:ChartFamily,subtype:ChartSubtype)=>string;
  onChoose:(choice:SpreadsheetDesktopChartTypeChoice)=>void;
}) {
  // Every row asks the catalogue whether its chart kind can be authored for
  // the current chart, which builds a chart plan per subtype. The closed menu
  // mounts none of these rows, so the catalogue is asked only once it opens.
  const [open, setOpen] = useState(false);
  const groups: OfficeChartTypeGroup[] = GROUPS.map(group=>({id:group.id,label:translate(group.labelKey),icon:group.icon,families:group.families}));
  return <DropdownMenu onOpenChange={setOpen}>
    <DropdownMenuTrigger asChild><OfficeHomeMenuTrigger variant="large" hasMenuIndicator
      disabled={disabledReason!==null} disabledReason={disabledReason??undefined}
      icon={<OfficeSemanticIcon semanticId="chart.type.column" size={26}/>}
      label={translate("spreadsheet.helperCopy.chartType.title")} testId="spreadsheet-ribbon-chart-change-type"/></DropdownMenuTrigger>
    <RibbonDropdownMenuContent align="start" className="docs-cupertino-menu" style={{width:230}} data-testid="spreadsheet-chart-type-menu">
      {open ? <OfficeChartTypeMenuItems groups={groups} disableUnavailableGroups={false}
        choice={(family,subtype)=>{const item=choice(family,subtype);return {enabled:item.enabled,reason:item.reason,checked:current.type===item.type&&current.subtype===subtype};}}
        subtypeLabel={label}
        onChoose={(family,subtype)=>onChoose(choice(family,subtype))}
        groupTestId={group=>`spreadsheet-chart-type-family-${group.id}`}
        itemTestId={(family,subtype)=>`spreadsheet-chart-type-${family}-${subtype}`}/> : null}
    </RibbonDropdownMenuContent>
  </DropdownMenu>;
}
