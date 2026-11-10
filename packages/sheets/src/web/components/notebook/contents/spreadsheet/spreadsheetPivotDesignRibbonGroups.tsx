"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Rows3, SquareSigma, LayoutGrid, BetweenHorizontalStart } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeStyleGallery, type OfficeStyleGalleryGroup, type OfficeStyleGalleryItem } from "@mrnote/core/web/components/office-ribbon/OfficeStyleGallery.tsx";
import spreadsheetHomeStyles from "./SpreadsheetHomeRibbon.module.css";
import { type PivotSubtotalPosition, PIVOT_SUBTOTAL_POSITIONS, type PivotGrandTotalMode, PIVOT_GRAND_TOTAL_MODES, type PivotReportLayout, PIVOT_REPORT_LAYOUTS, type PivotConfig } from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import { DropdownMenuRadioGroup, DropdownMenuSeparator } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { SpreadsheetRibbonMenuAction, DropdownMenuRadioItem } from "./SpreadsheetRibbonControls";
import { type ResolvedSpreadsheetVisibleControl } from "../../../../lib/spreadsheet/visible-command-authority";
import { type JSX } from "react";

/** What spreadsheetPivotDesignRibbonGroups reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetPivotDesignRibbonGroupsInputs {
  executePivotDesignMutation: (patch: Partial<PivotConfig>) => Promise<boolean>;
  pivotConfig: PivotConfig | null;
  pivotContextActive: boolean;
  pivotDesignCommandState: ResolvedSpreadsheetVisibleControl;
  pivotStyleGroups: readonly OfficeStyleGalleryGroup<string>[];
  pivotStyleItems: readonly OfficeStyleGalleryItem<string>[];
  pivotStyleQuickIds: string[];
  renderPivotStylePreview: (item: OfficeStyleGalleryItem<string>) => JSX.Element;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  visibleCommandDisclosureProps: (controlId: string) => { controlId: string; disabled: boolean; disabledReason: string | undefined; hidden: boolean; };
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function spreadsheetPivotDesignRibbonGroups({
  executePivotDesignMutation,
  pivotConfig,
  pivotContextActive,
  pivotDesignCommandState,
  pivotStyleGroups,
  pivotStyleItems,
  pivotStyleQuickIds,
  renderPivotStylePreview,
  t,
  visibleCommandDisclosureProps,
}: SpreadsheetPivotDesignRibbonGroupsInputs) {
  const pivotDesignRibbonGroups =
    pivotContextActive && pivotConfig
      ? [
          {
            id: "pivot-design-layout",
            label: t("toolbar.layoutMenu"),
            content: (
              <>
                <SpreadsheetRibbonMenuAction
                  {...visibleCommandDisclosureProps(
                    "XLSCTL.CONTEXT.PIVOT.DESIGN.SUBTOTALS_MENU",
                  )}
                  controlOccurrenceId="subtotals-menu"
                  controlSurface="pivot-design"
                  homeVariant="large"
                  icon={<Rows3 size={24} />}
                  label={t("spreadsheet.uiCopy.subtotals")}
                  testId="spreadsheet-ribbon-pivot-design-subtotals"
                  menuClassName="docs-cupertino-menu w-max max-w-[310px]"
                >
                  <DropdownMenuRadioGroup
                    value={
                      pivotConfig.subtotalPosition ??
                      (pivotConfig.rowSubtotals ? "bottom" : "none")
                    }
                    onValueChange={(rawValue) => {
                      const value = rawValue as PivotSubtotalPosition;
                      if (PIVOT_SUBTOTAL_POSITIONS.includes(value)) {
                        // Design > Subtotals sets every field's subtotals,
                        // a field's own (Field Settings) included.
                        void executePivotDesignMutation({
                          subtotalPosition: value,
                          rowSubtotals: value !== "none",
                          fieldSubtotals: undefined,
                        });
                      }
                    }}
                  >
                    {PIVOT_SUBTOTAL_POSITIONS.map((value) => (
                      <DropdownMenuRadioItem
                        key={value}
                        value={value}
                        data-control-id="XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"
                        data-control-occurrence-id={`subtotals-${value}`}
                        data-control-surface="pivot-design"
                        data-testid={`spreadsheet-ribbon-pivot-design-subtotals-${value}`}
                      >
                        {value === "none"
                          ? t("spreadsheet.uiCopy.doNotShowSubtotals")
                          : value === "bottom"
                            ? t("spreadsheet.uiCopy.showAllSubtotalsAtBottomOfGroup")
                            : t("spreadsheet.uiCopy.showAllSubtotalsAtTopOfGroup")}
                      </DropdownMenuRadioItem>
                    ))}
                  </DropdownMenuRadioGroup>
                </SpreadsheetRibbonMenuAction>

                <SpreadsheetRibbonMenuAction
                  {...visibleCommandDisclosureProps(
                    "XLSCTL.CONTEXT.PIVOT.DESIGN.GRAND_TOTALS_MENU",
                  )}
                  controlOccurrenceId="grand-totals-menu"
                  controlSurface="pivot-design"
                  homeVariant="large"
                  icon={<SquareSigma size={24} />}
                  label={t("spreadsheet.uiCopy.grandTotals")}
                  testId="spreadsheet-ribbon-pivot-design-grand-totals"
                  menuClassName="docs-cupertino-menu w-max max-w-[270px]"
                >
                  <DropdownMenuRadioGroup
                    value={pivotConfig.grandTotals ?? "both"}
                    onValueChange={(rawValue) => {
                      const value = rawValue as PivotGrandTotalMode;
                      if (PIVOT_GRAND_TOTAL_MODES.includes(value)) {
                        void executePivotDesignMutation({ grandTotals: value });
                      }
                    }}
                  >
                    {PIVOT_GRAND_TOTAL_MODES.map((value) => (
                      <DropdownMenuRadioItem
                        key={value}
                        value={value}
                        data-control-id="XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"
                        data-control-occurrence-id={`grand-totals-${value}`}
                        data-control-surface="pivot-design"
                        data-testid={`spreadsheet-ribbon-pivot-design-grand-totals-${value}`}
                      >
                        {value === "both"
                          ? t("spreadsheet.uiCopy.onForRowsAndColumns")
                          : value === "rows"
                            ? t("spreadsheet.uiCopy.onForRowsOnly")
                            : value === "columns"
                              ? t("spreadsheet.uiCopy.onForColumnsOnly")
                              : t("spreadsheet.uiCopy.offForRowsAndColumns")}
                      </DropdownMenuRadioItem>
                    ))}
                  </DropdownMenuRadioGroup>
                </SpreadsheetRibbonMenuAction>

                <SpreadsheetRibbonMenuAction
                  {...visibleCommandDisclosureProps(
                    "XLSCTL.CONTEXT.PIVOT.DESIGN.REPORT_LAYOUT_MENU",
                  )}
                  controlOccurrenceId="report-layout-menu"
                  controlSurface="pivot-design"
                  homeVariant="large"
                  icon={<LayoutGrid size={24} />}
                  label={t("spreadsheet.uiCopy.reportLayout")}
                  testId="spreadsheet-ribbon-pivot-design-report-layout"
                  menuClassName="docs-cupertino-menu w-max max-w-[290px]"
                >
                  <DropdownMenuRadioGroup
                    value={pivotConfig.reportLayout ?? "tabular"}
                    onValueChange={(rawValue) => {
                      const value = rawValue as PivotReportLayout;
                      if (PIVOT_REPORT_LAYOUTS.includes(value)) {
                        void executePivotDesignMutation({
                          reportLayout: value,
                          ...(value === "tabular"
                            ? {}
                            : { repeatItemLabels: false }),
                        });
                      }
                    }}
                  >
                    {PIVOT_REPORT_LAYOUTS.map((value) => (
                      <DropdownMenuRadioItem
                        key={value}
                        value={value}
                        data-control-id="XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"
                        data-control-occurrence-id={`report-layout-${value}`}
                        data-control-surface="pivot-design"
                        data-testid={`spreadsheet-ribbon-pivot-design-report-layout-${value}`}
                      >
                        {value === "compact"
                          ? t("spreadsheet.uiCopy.showInCompactForm")
                          : value === "outline"
                            ? t("spreadsheet.uiCopy.showInOutlineForm")
                            : t("spreadsheet.uiCopy.showInTabularForm")}
                      </DropdownMenuRadioItem>
                    ))}
                  </DropdownMenuRadioGroup>
                  <DropdownMenuSeparator />
                  <DropdownMenuRadioGroup
                    value={
                      pivotConfig.repeatItemLabels === true
                        ? "repeat"
                        : "do-not-repeat"
                    }
                    onValueChange={(value) =>
                      void executePivotDesignMutation({
                        repeatItemLabels: value === "repeat",
                      })
                    }
                  >
                    <DropdownMenuRadioItem
                      value="repeat"
                      disabled={
                        (pivotConfig.reportLayout ?? "tabular") !== "tabular"
                      }
                      data-control-id="XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"
                      data-control-occurrence-id="report-layout-repeat-labels"
                      data-control-surface="pivot-design"
                      data-testid="spreadsheet-ribbon-pivot-design-repeat-labels"
                    >
                      {t("spreadsheet.uiCopy.repeatAllItemLabels")}
                    </DropdownMenuRadioItem>
                    <DropdownMenuRadioItem
                      value="do-not-repeat"
                      data-control-id="XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"
                      data-control-occurrence-id="report-layout-do-not-repeat-labels"
                      data-control-surface="pivot-design"
                      data-testid="spreadsheet-ribbon-pivot-design-do-not-repeat-labels"
                    >
                      {t("spreadsheet.uiCopy.doNotRepeatItemLabels")}
                    </DropdownMenuRadioItem>
                  </DropdownMenuRadioGroup>
                </SpreadsheetRibbonMenuAction>

                <SpreadsheetRibbonMenuAction
                  {...visibleCommandDisclosureProps(
                    "XLSCTL.CONTEXT.PIVOT.DESIGN.BLANK_ROWS_MENU",
                  )}
                  controlOccurrenceId="blank-rows-menu"
                  controlSurface="pivot-design"
                  homeVariant="large"
                  icon={<BetweenHorizontalStart size={24} />}
                  label={t("spreadsheet.uiCopy.blankRows")}
                  testId="spreadsheet-ribbon-pivot-design-blank-rows"
                  menuClassName="docs-cupertino-menu w-max max-w-[290px]"
                >
                  <DropdownMenuRadioGroup
                    value={pivotConfig.blankRows === true ? "insert" : "remove"}
                    onValueChange={(value) =>
                      void executePivotDesignMutation({
                        blankRows: value === "insert",
                      })
                    }
                  >
                    <DropdownMenuRadioItem
                      value="insert"
                      data-control-id="XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"
                      data-control-occurrence-id="blank-rows-insert"
                      data-control-surface="pivot-design"
                      data-testid="spreadsheet-ribbon-pivot-design-blank-rows-insert"
                    >
                      {t("spreadsheet.uiCopy.insertBlankLineAfterEachItem")}
                    </DropdownMenuRadioItem>
                    <DropdownMenuRadioItem
                      value="remove"
                      data-control-id="XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"
                      data-control-occurrence-id="blank-rows-remove"
                      data-control-surface="pivot-design"
                      data-testid="spreadsheet-ribbon-pivot-design-blank-rows-remove"
                    >
                      {t("spreadsheet.uiCopy.removeBlankLineAfterEachItem")}
                    </DropdownMenuRadioItem>
                  </DropdownMenuRadioGroup>
                </SpreadsheetRibbonMenuAction>
              </>
            ),
          },
          {
            id: "pivot-design-style-options",
            label: t("spreadsheet.uiCopy.pivotTableStyleOptions"),
            content: (
              <div className="grid min-h-[64px] grid-cols-2 gap-x-3 px-2 text-[10px]">
                {([[
                  "showRowHeaders", t("spreadsheet.uiCopy.rowHeaders")],
                  ["showColumnHeaders", t("spreadsheet.uiCopy.columnHeaders")],
                  ["bandedRows", t("spreadsheet.uiCopy.bandedRows")],
                  ["bandedColumns", t("spreadsheet.uiCopy.bandedColumns")],
                ] as const).map(([key, label]) => <label key={key} className="flex items-center gap-1"><input type="checkbox" data-testid={`spreadsheet-ribbon-pivot-design-${key}`} checked={key.startsWith("show") ? pivotConfig[key] !== false : pivotConfig[key] === true} disabled={!pivotDesignCommandState.enabled} onChange={(event) => void executePivotDesignMutation({ [key]: event.target.checked })} />{label}</label>)}
              </div>
            ),
          },
          {
            id: "pivot-design-styles",
            label: t("spreadsheet.uiCopy.pivotTableStyles"),
            content: (
              <div data-control-id="XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY" data-control-surface="pivot-design" data-testid="spreadsheet-ribbon-pivot-design-style-authority">
                <OfficeStyleGallery<string> application="spreadsheet" ariaLabel={t("spreadsheet.uiCopy.pivotTableStyles")} className={`${spreadsheetHomeStyles.contextualStyleGallery} ${spreadsheetHomeStyles.pivotStyleGallery}`} disabled={!pivotDesignCommandState.enabled} groups={pivotStyleGroups} groupHeadings items={pivotStyleItems} menuGrid={{ columns: 7, thumbnail: [60, 44] }} menuMode="preview-only" labels={{ more: t("spreadsheet.uiCopy.moreStyles") }} quickIds={pivotStyleQuickIds} renderPreview={renderPivotStylePreview} testIdPrefix="spreadsheet-pivot-style-gallery" value={pivotConfig.styleName ?? pivotConfig.native?.styleName ?? "PivotStyleLight16"} onPreviewChange={() => undefined} onCommit={(styleName) => void executePivotDesignMutation({ styleName })} />
              </div>
            ),
          },
        ]
      : [];
  return {
    pivotDesignRibbonGroups,
  };
}
