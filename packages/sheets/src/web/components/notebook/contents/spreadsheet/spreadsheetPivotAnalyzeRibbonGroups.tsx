"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type RefObject } from "react";
import { RefreshCw, TableProperties, BoxSelect, Table2, Hash, Tags, Eraser, RotateCcw, FilterX, Trash2, Locate, ChartNoAxesCombined, ListTree, SquarePlus } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { SpreadsheetInsertIcon } from "./SpreadsheetInsertRibbonIcons";
import { SPREADSHEET_PIVOT_CHART_TYPES, type SpreadsheetPivotChartType } from "@mrnote/core/web/lib/spreadsheet/pivot-chart.ts";
import { type SpreadsheetChartType } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { type PivotConfig } from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import { DropdownMenuSeparator } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { spreadsheetChartTypeLabel } from "./SpreadsheetChartPanel";
import { SpreadsheetRibbonAction, SpreadsheetRibbonMenuAction, DropdownMenuItem } from "./SpreadsheetRibbonControls";
import { type ResolvedSpreadsheetVisibleControl } from "../../../../lib/spreadsheet/visible-command-authority";

/** What spreadsheetPivotAnalyzeRibbonGroups reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetPivotAnalyzeRibbonGroupsInputs {
  clearPivot: () => Promise<boolean>;
  clearPivotAll: () => Promise<boolean>;
  clearPivotFilters: () => Promise<boolean>;
  executePivotChartCreate: (chartType: SpreadsheetPivotChartType) => Promise<boolean>;
  executePivotRefresh: () => Promise<boolean>;
  executeVisibleCommand: (controlId: string) => Promise<boolean>;
  insertChartIconName: (type: SpreadsheetChartType) => "chartColumn" | "chartLine" | "chartScatter" | "chartPie" | "chartBar" | "chartStatistical" | "chartCombo" | "chartOther";
  openPivotSlicerInsertDialog: (controlId: "XLSCTL.INSERT.FILTER.SLICER" | "XLSCTL.CONTEXT.PIVOT.INSERT_SLICER") => Promise<boolean>;
  pivotChartCreateReason: string | undefined;
  pivotChartCreateState: ResolvedSpreadsheetVisibleControl;
  pivotChartOpen: boolean;
  pivotConfig: PivotConfig | null;
  pivotMoveDialogReturnFocusRef: RefObject<HTMLButtonElement | null>;
  pivotNativeSourceDirty: boolean;
  pivotOrChartContextActive: boolean;
  pivotPanelOpen: boolean;
  pivotSourceDialogReturnFocusRef: RefObject<HTMLButtonElement | null>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  togglePivotShowDrillButtons: () => Promise<boolean>;
  visibleCommandActionProps: (controlId: string) => { controlId: string; disabled: boolean; disabledReason: string | undefined; hidden: boolean; onClick: () => Promise<boolean>; };
  visibleCommandDisclosureProps: (controlId: string) => { controlId: string; disabled: boolean; disabledReason: string | undefined; hidden: boolean; };
  visibleCommandMenuProps: (controlId: string, accessibleLabel?: string, selectionReason?: string) => { "data-control-id": string; "aria-label": string | undefined; disabled: boolean; hidden: boolean; title: string | undefined; };
  visibleCommandState: (controlId: string) => ResolvedSpreadsheetVisibleControl;
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function spreadsheetPivotAnalyzeRibbonGroups({
  clearPivot,
  clearPivotAll,
  clearPivotFilters,
  executePivotChartCreate,
  executePivotRefresh,
  executeVisibleCommand,
  insertChartIconName,
  openPivotSlicerInsertDialog,
  pivotChartCreateReason,
  pivotChartCreateState,
  pivotChartOpen,
  pivotConfig,
  pivotMoveDialogReturnFocusRef,
  pivotNativeSourceDirty,
  pivotOrChartContextActive,
  pivotPanelOpen,
  pivotSourceDialogReturnFocusRef,
  t,
  togglePivotShowDrillButtons,
  visibleCommandActionProps,
  visibleCommandDisclosureProps,
  visibleCommandMenuProps,
  visibleCommandState,
}: SpreadsheetPivotAnalyzeRibbonGroupsInputs) {
  const pivotAnalyzeRibbonGroups =
    pivotOrChartContextActive && pivotConfig
      ? [
          {
            id: "pivot-analyze-data",
            label: t("spreadsheet.chartFormatPane.dataTab"),
            content: (
              <>
                <SpreadsheetRibbonAction
                  {...visibleCommandActionProps(
                    pivotConfig.native || pivotConfig.cacheSnapshot
                      ? "XLSCTL.CONTEXT.PIVOT.REFRESH.MRAI.APPLY"
                      : "XLSCTL.CONTEXT.PIVOT.REFRESH",
                  )}
                  controlOccurrenceId="refresh"
                  controlSurface="pivot-analyze"
                  active={pivotNativeSourceDirty}
                  homeVariant="large"
                  icon={<RefreshCw size={24} />}
                  label={
                    t(pivotNativeSourceDirty ? "spreadsheet.dynamicCopy.pivotRefreshChanged" : "spreadsheet.dynamicCopy.pivotRefresh")
                  }
                  onClick={() => void executePivotRefresh()}
                  testId="spreadsheet-ribbon-pivot-refresh"
                />
                <SpreadsheetRibbonAction
                  {...visibleCommandActionProps(
                    "XLSCTL.CONTEXT.PIVOT.CHANGE_SOURCE",
                  )}
                  buttonRef={pivotSourceDialogReturnFocusRef}
                  controlOccurrenceId="change-source"
                  controlSurface="pivot-analyze"
                  homeVariant="large"
                  icon={<TableProperties size={24} />}
                  label={
                    t("spreadsheet.uiCopy.changeDataSource")
                  }
                  onClick={() =>
                    executeVisibleCommand(
                      "XLSCTL.CONTEXT.PIVOT.CHANGE_SOURCE",
                    )
                  }
                  testId="spreadsheet-ribbon-pivot-source"
                />
              </>
            ),
          },
          {
            id: "pivot-analyze-actions",
            label: t("spreadsheet.uiCopy.actions"),
            content: (
              <>
                <SpreadsheetRibbonMenuAction
                  {...visibleCommandDisclosureProps(
                    "XLSCTL.CONTEXT.PIVOT.SELECT_MENU",
                  )}
                  controlOccurrenceId="select-menu"
                  controlSurface="pivot-analyze"
                  homeVariant="large"
                  icon={<BoxSelect size={24} />}
                  label={t("presentation.tableRibbon.select")}
                  testId="spreadsheet-ribbon-pivot-select"
                  menuClassName="docs-cupertino-menu w-max max-w-[230px]"
                >
                  {/* Excel's order: Labels and Values, Values, Labels, Entire PivotTable. */}
                  <DropdownMenuItem
                    data-control-id="XLSCTL.CONTEXT.PIVOT.SELECT_LABELS_VALUES.MRAI"
                    data-testid="spreadsheet-ribbon-pivot-select-labels-values"
                    onSelect={() =>
                      void executeVisibleCommand(
                        "XLSCTL.CONTEXT.PIVOT.SELECT_LABELS_VALUES.MRAI",
                      )
                    }
                  >
                    <Table2 aria-hidden="true" size={16} />
                    {t("spreadsheet.uiCopy.pivotLabelsAndValues")}
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    data-control-id="XLSCTL.CONTEXT.PIVOT.SELECT_VALUES.MRAI"
                    data-testid="spreadsheet-ribbon-pivot-select-values"
                    onSelect={() =>
                      void executeVisibleCommand(
                        "XLSCTL.CONTEXT.PIVOT.SELECT_VALUES.MRAI",
                      )
                    }
                  >
                    <Hash aria-hidden="true" size={16} />
                    {t("spreadsheet.pivotValues")}
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    data-control-id="XLSCTL.CONTEXT.PIVOT.SELECT_LABELS.MRAI"
                    data-testid="spreadsheet-ribbon-pivot-select-labels"
                    onSelect={() =>
                      void executeVisibleCommand(
                        "XLSCTL.CONTEXT.PIVOT.SELECT_LABELS.MRAI",
                      )
                    }
                  >
                    <Tags aria-hidden="true" size={16} />
                    {t("spreadsheet.uiCopy.pivotFieldLabels")}
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    data-control-id="XLSCTL.CONTEXT.PIVOT.SELECT_ENTIRE.MRAI"
                    data-testid="spreadsheet-ribbon-pivot-select-entire"
                    onSelect={() =>
                      void executeVisibleCommand(
                        "XLSCTL.CONTEXT.PIVOT.SELECT_ENTIRE.MRAI",
                      )
                    }
                  >
                    <BoxSelect aria-hidden="true" size={16} />
                    {t("spreadsheet.uiCopy.entirePivotTable")}
                  </DropdownMenuItem>
                </SpreadsheetRibbonMenuAction>
                <SpreadsheetRibbonMenuAction
                  {...visibleCommandDisclosureProps(
                    "XLSCTL.CONTEXT.PIVOT.CLEAR_MENU",
                  )}
                  controlOccurrenceId="clear-menu"
                  controlSurface="pivot-analyze"
                  homeVariant="large"
                  icon={<Eraser size={24} />}
                  label={t("spreadsheet.clearMenu")}
                  testId="spreadsheet-ribbon-pivot-clear"
                  menuClassName="docs-cupertino-menu w-max max-w-[240px]"
                >
                  <DropdownMenuItem
                    {...visibleCommandMenuProps(
                      "XLSCTL.CONTEXT.PIVOT.CLEAR_ALL.MRAI.APPLY",
                      t("spreadsheet.uiCopy.clearAll"),
                    )}
                    data-testid="spreadsheet-ribbon-pivot-clear-all"
                    onSelect={() => void clearPivotAll()}
                  >
                    <RotateCcw aria-hidden="true" size={16} />
                    {t("spreadsheet.uiCopy.clearAll")}
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    data-control-id="XLSCTL.CONTEXT.PIVOT.CLEAR_FILTERS.MRAI.APPLY"
                    data-testid="spreadsheet-ribbon-pivot-clear-filters"
                    disabled={
                      !visibleCommandState(
                        "XLSCTL.CONTEXT.PIVOT.CLEAR_FILTERS.MRAI.APPLY",
                      ).enabled ||
                      !pivotConfig.filters.some(
                        (filter) => filter.selected.length > 0,
                      )
                    }
                    onSelect={() => void clearPivotFilters()}
                  >
                    <FilterX aria-hidden="true" size={16} />
                    {t("spreadsheet.uiCopy.clearFilters")}
                  </DropdownMenuItem>
                  <DropdownMenuSeparator />
                  <DropdownMenuItem
                    data-control-id="XLSCTL.CONTEXT.PIVOT.CLEAR.MRAI.APPLY"
                    data-testid="spreadsheet-ribbon-pivot-remove"
                    onSelect={() => void clearPivot()}
                  >
                    <Trash2 aria-hidden="true" size={16} />
                    {t("spreadsheet.uiCopy.removePivotTable")}
                  </DropdownMenuItem>
                </SpreadsheetRibbonMenuAction>
                <SpreadsheetRibbonAction
                  {...visibleCommandActionProps("XLSCTL.CONTEXT.PIVOT.MOVE")}
                  buttonRef={pivotMoveDialogReturnFocusRef}
                  controlOccurrenceId="move"
                  controlSurface="pivot-analyze"
                  homeVariant="large"
                  icon={<Locate size={24} />}
                  label={t("spreadsheet.uiCopy.movePivotTable")}
                  onClick={() =>
                    void executeVisibleCommand("XLSCTL.CONTEXT.PIVOT.MOVE")
                  }
                  testId="spreadsheet-ribbon-pivot-move"
                />
              </>
            ),
          },
          {
            id: "pivot-analyze-tools",
            label: t("spreadsheet.ribbon.toolsGroup"),
            content: (
              <>
                <SpreadsheetRibbonAction
                  {...visibleCommandActionProps(
                    "XLSCTL.CONTEXT.PIVOT.CHART_PREVIEW.MRAI",
                  )}
                  active={pivotChartOpen}
                  controlOccurrenceId="chart-preview"
                  controlSurface="pivot-analyze"
                  homeVariant="large"
                  icon={<ChartNoAxesCombined size={24} />}
                  label={
                    t(pivotChartOpen ? "spreadsheet.dynamicCopy.hidePivotChartPreview" : "spreadsheet.dynamicCopy.pivotChartPreview")
                  }
                  onClick={() =>
                    executeVisibleCommand(
                      "XLSCTL.CONTEXT.PIVOT.CHART_PREVIEW.MRAI",
                    )
                  }
                  testId="spreadsheet-ribbon-pivot-chart"
                />
                <SpreadsheetRibbonMenuAction
                  {...visibleCommandDisclosureProps(
                    "XLSCTL.CONTEXT.PIVOT.CHART_MENU",
                  )}
                  controlOccurrenceId="chart-menu"
                  controlSurface="pivot-analyze"
                  homeVariant="large"
                  icon={<ChartNoAxesCombined size={24} />}
                  label={t("spreadsheet.uiCopy.pivotChart")}
                  testId="spreadsheet-ribbon-pivot-chart-create"
                  menuClassName="docs-cupertino-menu w-max max-w-[260px]"
                >
                  {SPREADSHEET_PIVOT_CHART_TYPES.map((type) => {
                    const label = spreadsheetChartTypeLabel(t, type);
                    const accessibleLabel = pivotChartCreateReason
                      ? `${label}: ${pivotChartCreateReason}`
                      : label;
                    return (
                      <DropdownMenuItem
                        key={type}
                        aria-label={accessibleLabel}
                        data-control-id="XLSCTL.CONTEXT.PIVOT.CHART.MRAI.APPLY"
                        data-control-occurrence-id={`chart-${type}`}
                        data-control-surface="pivot-analyze"
                        data-testid={`spreadsheet-ribbon-pivot-chart-type-${type}`}
                        disabled={!pivotChartCreateState.enabled}
                        title={accessibleLabel}
                        onSelect={() => void executePivotChartCreate(type)}
                      >
                        <SpreadsheetInsertIcon
                          name={insertChartIconName(type)}
                          size={16}
                        />
                        {label}
                      </DropdownMenuItem>
                    );
                  })}
                </SpreadsheetRibbonMenuAction>
              </>
            ),
          },
          {
            id: "pivot-analyze-filter",
            label: t("spreadsheet.uiCopy.pivotAnalyzeFilterGroup"),
            content: (
              <SpreadsheetRibbonAction
                {...visibleCommandActionProps(
                  "XLSCTL.CONTEXT.PIVOT.INSERT_SLICER",
                )}
                controlOccurrenceId="insert-slicer"
                controlSurface="pivot-analyze"
                homeVariant="large"
                icon={<OfficeSemanticIcon semanticId="table.slicer" size={27} />}
                label={t("spreadsheet.uiCopy.insertSlicer")}
                onClick={() =>
                  void openPivotSlicerInsertDialog(
                    "XLSCTL.CONTEXT.PIVOT.INSERT_SLICER",
                  )
                }
                testId="spreadsheet-ribbon-pivot-insert-slicer"
              />
            ),
          },
          {
            id: "pivot-analyze-show",
            label: t("spreadsheet.view.show"),
            content: (
              <>
                <SpreadsheetRibbonAction
                  {...visibleCommandActionProps(
                    "XLSCTL.CONTEXT.PIVOT.FIELD_LIST",
                  )}
                  active={pivotPanelOpen}
                  controlOccurrenceId="field-list"
                  controlSurface="pivot-analyze"
                  homeVariant="large"
                  icon={<ListTree size={24} />}
                  label={t("spreadsheet.uiCopy.fieldList")}
                  onClick={() =>
                    executeVisibleCommand("XLSCTL.CONTEXT.PIVOT.FIELD_LIST")
                  }
                  testId="spreadsheet-ribbon-pivot-fields"
                />
                <SpreadsheetRibbonAction
                  {...visibleCommandActionProps(
                    "XLSCTL.CONTEXT.PIVOT.SHOW_EXPAND_COLLAPSE_BUTTONS.MRAI.APPLY",
                  )}
                  active={pivotConfig.showDrillButtons !== false}
                  controlOccurrenceId="expand-collapse-buttons"
                  controlSurface="pivot-analyze"
                  homeVariant="large"
                  icon={<SquarePlus size={24} />}
                  label={t("spreadsheet.uiCopy.buttons")}
                  onClick={() => void togglePivotShowDrillButtons()}
                  testId="spreadsheet-ribbon-pivot-expand-collapse-buttons"
                />
              </>
            ),
          },
        ]
      : [];
  return {
    pivotAnalyzeRibbonGroups,
  };
}
