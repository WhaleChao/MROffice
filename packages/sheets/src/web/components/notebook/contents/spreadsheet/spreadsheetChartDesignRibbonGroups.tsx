"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { sheetChartPlotsByRows, type ChartQuickLayoutId } from "../../../../lib/spreadsheet/drawing-buttons";
import { type RefObject, type ReactNode } from "react";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import spreadsheetHomeStyles from "./SpreadsheetHomeRibbon.module.css";
import { type DocumentChartAreaStyleKey } from "@mrnote/core/web/components/console/editor/native-document/DocumentChartFormatRibbon.tsx";
import { DocumentChartStyleGallery } from "@mrnote/docs/web/components/console/editor/native-document/DocumentChartStyleGallery.tsx";
import { documentChartStyleLabel, type DocumentChartStyleId } from "@mrnote/core/web/components/console/editor/native-document/documentChartStyles.ts";
import { spreadsheetChartSupportsMinorGridlines } from "@mrnote/core/web/lib/spreadsheet/chart-v2-binding.ts";
import { chartDataLabelPlacements, chartElementAvailability, chartDataLabelPlacement, type ChartElementIntent, chartWithTrendlineOption } from "@mrnote/core/web/lib/office-charts/chart-elements.ts";
import type { ChartTrendline } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { OfficeChartQuickLayoutGallery } from "@mrnote/core/web/components/office-ribbon/OfficeChartQuickLayoutPreview.tsx";
import { type OfficeChartStylesSlot, OfficeChartQuickButtons, officeChartFilterNames } from "@mrnote/core/web/components/office-ribbon/OfficeChartQuickButtons.tsx";
import { officeChartFilterControls } from "@mrnote/core/web/lib/office-charts/chart-filters.ts";
import { type OfficeChartElementEntryState, type OfficeChartElementEntryKey, officeChartElementView, type OfficeChartElementView, type OfficeChartLegendPosition, OfficeChartElementMenuItems, OfficeChartDesignRibbon, officeChartHiddenEmptyCellsMenu } from "@mrnote/core/web/components/office-ribbon/OfficeChartDesignRibbon.tsx";
import { OfficeChartColorPaletteMenuItems } from "@mrnote/core/web/components/office-ribbon/OfficeChartColorPaletteMenu.tsx";
import { normalizeSpreadsheetChartFormat, type SpreadsheetChartFormat } from "@mrnote/core/web/lib/spreadsheet/chart-format.ts";
import type { ChartBlockAttrsV2 } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { type ChartFamily, type ChartSubtype } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { SpreadsheetChartTypeMenu, type SpreadsheetDesktopChartTypeChoice } from "./SpreadsheetChartTypeMenu";
import { importedChartOnSharedModel } from "./SpreadsheetChartPanel";
import { type SpreadsheetVisibleCommandReason } from "../../../../lib/spreadsheet/visible-command-authority";
import { type SpreadsheetRibbonTabId } from "../../../../lib/spreadsheet/SpreadsheetRibbon";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { normalizeChartNativeLook } from "@mrnote/core/web/lib/office-chart-native-look.ts";
import { normalizeTrendline, normalizeErrorBars, DEFAULT_TRENDLINE_PERIOD, type SpreadsheetChart, type SpreadsheetChartType, type SpreadsheetChartColorPaletteId } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { type SpreadsheetObject } from "./spreadsheetSheetModel";
import { OfficeHomeMenuTrigger } from "./SpreadsheetRibbonControls";
import { type ChartDesignDialogAuthority, type ChartDesignMutationReceipt } from "./spreadsheetMutationReceipts";
import { SPREADSHEET_CHART_LEGEND_LABEL_KEYS, spreadsheetChartDesignTestId } from "./spreadsheetChartInsertMenus";
import { type ResolvedSpreadsheetVisibleControl } from "../../../../lib/spreadsheet/visible-command-authority";
import { type ChartBlockFormat, type ChartData } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { type ChartStylePresetId } from "@mrnote/core/web/lib/office-charts/chart-style-presets.generated.ts";
import { type SpreadsheetChartLegendPosition } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { type Dispatch, type SetStateAction } from "react";

/** What spreadsheetChartDesignRibbonGroups reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetChartDesignRibbonGroupsInputs {
  chartColorCommandState: ResolvedSpreadsheetVisibleControl;
  chartColorDisabledReason: string | undefined;
  chartDataReturnFocusRef: RefObject<HTMLButtonElement | null>;
  chartDesignActionProps: (controlId: string) => { controlId: string; disabled: boolean; disabledReason: string | undefined; hidden: boolean; };
  chartDesignControlState: (controlId: string) => { state: { enabled: boolean; visible: boolean; reasonKey: SpreadsheetVisibleCommandReason | null; featureId: string; controlId: string; }; reason: string | undefined; dom: { "data-control-id": string; "data-control-surface": string; disabled: boolean; hidden: boolean; title: string | undefined; }; };
  chartFormatFocusTokenRef: RefObject<number>;
  chartMoveTriggerRef: RefObject<HTMLButtonElement | null>;
  chartQuickLayoutCalibrated: boolean;
  chartQuickLayoutCommandState: ResolvedSpreadsheetVisibleControl;
  chartQuickLayoutDisabledReason: string | undefined;
  chartQuickLayoutRefusal: (layoutId: ChartQuickLayoutId) => string | null;
  chartStyleCommandState: ResolvedSpreadsheetVisibleControl;
  chartStyleDisabledReason: string | undefined;
  chooseDesktopChartType: (choice: SpreadsheetDesktopChartTypeChoice) => void;
  desktopChartAttrs: { title: string; schemaVersion: 2; chartType: ChartFamily; chartSubtype: ChartSubtype; axisTitleX?: string; axisTitleY?: string; data: ChartData; stylePreset: ChartStylePresetId; format?: ChartBlockFormat; } | null;
  desktopChartFormatReason: string | null;
  desktopChartModelReason: string | null;
  desktopChartPreviewAttrs: ChartBlockAttrsV2 | null;
  desktopChartStyleCommit: (id: DocumentChartStyleId, authority?: ChartDesignDialogAuthority | null) => boolean;
  desktopChartStyleFormat: (id: DocumentChartStyleId) => SpreadsheetChartFormat | null;
  desktopTypeChoice: (family: ChartFamily, subtype: ChartSubtype) => SpreadsheetDesktopChartTypeChoice;
  executeChartDesignMutation: (controlId: ChartDesignMutationReceipt["controlId"], parameter: string) => Promise<boolean>;
  executeVisibleCommand: (controlId: string) => Promise<boolean>;
  locale: string;
  openDesktopChartMove: () => void;
  patchDesktopChartFilters: (picked: ChartBlockAttrsV2) => boolean;
  selectedChartColorPaletteId: SpreadsheetChartColorPaletteId | undefined;
  selectedChartDesignKey: string | null;
  selectedChartDesignLegendPosition: SpreadsheetChartLegendPosition;
  selectedChartDesignShowAxisLabels: boolean;
  selectedChartDesignShowDataLabels: boolean;
  selectedChartDesignShowDataTable: boolean;
  selectedChartDesignShowGridlines: boolean;
  selectedChartDesignShowLegend: boolean;
  selectedChartDesignTarget: { readonly kind: "drawing"; readonly chart: SpreadsheetObject; } | { readonly kind: "native"; readonly chart: SpreadsheetChart; } | null;
  selectedChartDesignTitle: string;
  selectedChartDesignType: SpreadsheetChartType | undefined;
  selectedChartModel: ChartBlockAttrsV2 | undefined;
  selectedChartQuickLayoutId: "layout-1" | "layout-2" | "layout-3" | "layout-4" | "layout-5" | "layout-6" | "layout-7" | "layout-8" | "layout-9" | "layout-10" | "layout-11" | null;
  setChartColorPalettePreview: Dispatch<SetStateAction<{ chartKey: string; paletteId: SpreadsheetChartColorPaletteId; } | null>>;
  setChartDesktopPreview: Dispatch<SetStateAction<{ key: string; source: SpreadsheetChart | SpreadsheetObject; styleId: DocumentChartStyleId; format: SpreadsheetChartFormat | null; } | null>>;
  setChartFormatFocus: Dispatch<SetStateAction<{ area: DocumentChartAreaStyleKey; token: number; } | null>>;
  setChartQuickLayoutPreview: Dispatch<SetStateAction<{ chartKey: string; layoutId: ChartQuickLayoutId; } | null>>;
  setRibbonTab: (next: SpreadsheetRibbonTabId) => void;
  sheetChartEditReason: (chart: SpreadsheetChart | SpreadsheetObject) => string | null;
  sheetChartQuick: { key: string | null; node: ReactNode; };
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  visibleCommandDisclosureProps: (controlId: string) => { controlId: string; disabled: boolean; disabledReason: string | undefined; hidden: boolean; };
  visibleCommandReason: (reasonKey: SpreadsheetVisibleCommandReason | null) => string | undefined;
  workbookThemeAccentColors: readonly string[];
  workbookTopologyLocked: boolean;
  /** The selected chart's size on the sheet (px): Chart Styles' tiles are
   *  that chart drawn smaller, laid out at its own size. */
  selectedChartStyleGallerySize?: { readonly width: number; readonly height: number };
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function spreadsheetChartDesignRibbonGroups({
  chartColorCommandState,
  chartColorDisabledReason,
  chartDataReturnFocusRef,
  chartDesignActionProps,
  chartDesignControlState,
  chartFormatFocusTokenRef,
  chartMoveTriggerRef,
  chartQuickLayoutCalibrated,
  chartQuickLayoutCommandState,
  chartQuickLayoutDisabledReason,
  chartQuickLayoutRefusal,
  chartStyleCommandState,
  chartStyleDisabledReason,
  chooseDesktopChartType,
  desktopChartAttrs,
  desktopChartFormatReason,
  desktopChartModelReason,
  desktopChartPreviewAttrs,
  desktopChartStyleCommit,
  desktopChartStyleFormat,
  desktopTypeChoice,
  executeChartDesignMutation,
  executeVisibleCommand,
  locale,
  openDesktopChartMove,
  patchDesktopChartFilters,
  selectedChartColorPaletteId,
  selectedChartDesignKey,
  selectedChartDesignLegendPosition,
  selectedChartDesignShowAxisLabels,
  selectedChartDesignShowDataLabels,
  selectedChartDesignShowDataTable,
  selectedChartDesignShowGridlines,
  selectedChartDesignShowLegend,
  selectedChartDesignTarget,
  selectedChartDesignTitle,
  selectedChartDesignType,
  selectedChartModel,
  selectedChartQuickLayoutId,
  setChartColorPalettePreview,
  setChartDesktopPreview,
  setChartFormatFocus,
  setChartQuickLayoutPreview,
  setRibbonTab,
  sheetChartEditReason,
  sheetChartQuick,
  t,
  visibleCommandDisclosureProps,
  visibleCommandReason,
  workbookThemeAccentColors,
  workbookTopologyLocked,
  selectedChartStyleGallerySize,
}: SpreadsheetChartDesignRibbonGroupsInputs) {
  // Built with the page, not with the (deferred) ribbon: building it also
  // hands the selected chart its Chart Elements / Styles / Filters buttons
  // (sheetChartQuick), which the chart itself draws in this render.
  const chartDesignRibbonGroups = selectedChartDesignTarget
    ? (() => {
        const chartType = selectedChartDesignType ?? "column";
        const axisTitleAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.AXIS_TITLE_X.MRAI.APPLY",
        );
        const chartTitleAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.TITLE.MRAI.APPLY",
        );
        const legendAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.LEGEND_POSITION.MRAI.APPLY",
        );
        const legendToggleAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.LEGEND",
        );
        const dataLabelsAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.DATA_LABELS",
        );
        const dataLabelPositionAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.DATA_LABEL_POSITION.MRAI.APPLY",
        );
        // The Data Labels submenu: None, then either Office's
        // positions for this family (behind an Automatic row) or a
        // filled map's Map labels tri-state, which has no
        // automatic member. One list, from the shared table the
        // deck and Docs draw theirs from.
        const dataLabelPlacements = desktopChartAttrs
          ? chartDataLabelPlacements(desktopChartAttrs)
          : [];
        const dataLabelsAreMapLabels = desktopChartAttrs
          ? chartElementAvailability(desktopChartAttrs)
              .dataLabelRegionLayouts.length > 0
          : false;
        const selectedDataLabelPlacement = desktopChartAttrs
          ? chartDataLabelPlacement(desktopChartAttrs)
          : null;
        const dataTableAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.DATA_TABLE",
        );
        const axesAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.AXES",
        );
        const gridlinesAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.GRIDLINES",
        );
        const trendlineAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.TRENDLINE.MRAI.APPLY",
        );
        const errorBarsAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
        );
        const upDownBarsAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.UP_DOWN_BARS.MRAI.APPLY",
        );
        const linesMenuAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.LINES_MENU",
        );
        const linesWriterAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.LINES.MRAI.APPLY",
        );
        // Add Chart Element's rows speak for the sheet's own
        // command ids: each row's authority (enabled state, reason,
        // `data-control-id`) is the one the ribbon always asked.
        const axisTitleYAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.AXIS_TITLE_Y.MRAI.APPLY",
        );
        const axisTitleDepthAuthority = chartDesignControlState(
          "XLSCTL.CONTEXT.CHART.AXIS_TITLE_DEPTH.MRAI.APPLY",
        );
        const sheetChart = selectedChartDesignTarget.chart;
        // An imported drawing carries its analysis specs
        // unnormalized; read them as the chart renderer does.
        const sheetAnalysis = {
          trendline: normalizeTrendline(sheetChart.trendline),
          errorBars: normalizeErrorBars(sheetChart.errorBars),
        };
        const sheetChartFormat = sheetChart.chartFormat
          ? normalizeSpreadsheetChartFormat(sheetChart.chartFormat)
          : {};
        // A chart whose desktop format holds atoms this editor cannot
        // restate keeps it. An imported chart on the one model is edited by
        // the shared chart editor instead (each pick a model write, written
        // into its part), as Format and Chart Styles already treat it: the
        // flat reader's refusal (an openpyxl legend over the plot,
        // `xlsx_chart_layout_custom_legend`) is not this chart's.
        const sheetFormatReason =
          (!sheetChartFormat || sheetChart.nativeChartFormatBlocker) &&
          !importedChartOnSharedModel(selectedChartDesignTarget)
            ? (sheetChartEditReason(sheetChart) ?? t("spreadsheet.uiCopy.thisChartContainsNativeFormattingThatCannotBeEdited"))
            : undefined;
        const authorityEntry = (
          authority: ReturnType<typeof chartDesignControlState>,
          formatReason?: string,
        ): OfficeChartElementEntryState => ({
          disabled: !authority.state.enabled || Boolean(formatReason),
          reason: formatReason ?? authority.reason,
          dom: authority.dom,
        });
        const sheetElementEntry = (
          key: OfficeChartElementEntryKey,
          row?: string,
        ): OfficeChartElementEntryState | undefined => {
          switch (key) {
            case "chartTitle":
            case "chartTitleOption":
              return authorityEntry(chartTitleAuthority);
            case "axisTitles":
            case "axisTitleX":
              return authorityEntry(axisTitleAuthority);
            case "axisTitleY":
              return authorityEntry(axisTitleYAuthority);
            case "axisTitleDepth":
              return authorityEntry(axisTitleDepthAuthority);
            case "axes":
              return authorityEntry(axesAuthority);
            case "axisX":
            case "axisY":
              return authorityEntry(axesAuthority, sheetFormatReason);
            case "gridlines":
              return authorityEntry(gridlinesAuthority);
            case "gridlinesOption":
              return authorityEntry(gridlinesAuthority, sheetFormatReason);
            case "dataLabels":
            case "dataLabelsNone":
              return authorityEntry(dataLabelsAuthority);
            case "dataLabelPosition":
            case "dataLabelCallout":
              return authorityEntry(dataLabelPositionAuthority);
            case "dataTable":
              return authorityEntry(dataTableAuthority);
            case "dataTableOption":
              return authorityEntry(
                dataTableAuthority,
                row === "no-keys" ? sheetFormatReason : undefined,
              );
            case "errorBars":
              return authorityEntry(errorBarsAuthority);
            case "errorBarsOption":
              // Standard Error, Percentage and Standard Deviation
              // (`stdErr` / `percentage` / `stdDev`), as Excel.
              return authorityEntry(errorBarsAuthority);
            case "legend":
            case "legendNone":
              return authorityEntry(legendToggleAuthority);
            case "legendPosition":
              return authorityEntry(legendAuthority);
            case "lines":
              return authorityEntry(linesMenuAuthority);
            case "lineOption":
              return authorityEntry(linesWriterAuthority);
            case "trendline":
            case "trendlineOption":
              return authorityEntry(trendlineAuthority);
            case "upDownBars":
            case "upDownBarsOption":
              return authorityEntry(upDownBarsAuthority);
          }
        };
        const sheetBaseView = desktopChartAttrs
          ? officeChartElementView(desktopChartAttrs)
          : null;
        const sheetAxisShown = (key: "axisX" | "axisY") =>
          sheetChartFormat?.[key] ?? selectedChartDesignShowAxisLabels;
        const sheetGridShown = (key: "gridlinesX" | "gridlinesY") =>
          sheetChartFormat?.[key] ?? selectedChartDesignShowGridlines;
        const sheetErrorAmount = sheetAnalysis.errorBars?.type;
        const sheetElementView: OfficeChartElementView = {
          // The chart plan answers which elements the family takes;
          // a chart it cannot plan leaves that to the authority.
          available: {
            ...(sheetBaseView?.available ?? {
              axisX: true,
              axisY: true,
              axisSecondaryY: false,
              axisTitleX: true,
              axisTitleY: true,
              gridlines: true,
              dataLabels: true,
              dataTable: true,
              errorBars: true,
              dropLines: true,
              highLowLines: true,
              seriesLines: true,
              trendline: true,
              upDownBars: true,
            }),
            // Minor gridlines ride a calibrated column / bar
            // layout, and every other classic family states them
            // in its desktop format (the Gridlines command writes
            // both); only an Office 2016 chart has no writer.
            minorGridlines:
              chartQuickLayoutCalibrated ||
              spreadsheetChartSupportsMinorGridlines({
                type: selectedChartDesignType ?? "column",
                chartSubtype: sheetChart.chartSubtype,
              }),
            // Axis Titles > Depth: a chart with a depth (series) axis
            // (the plan's own answer; a chart it cannot plan has none).
            axisTitleDepth: sheetBaseView?.available.axisTitleDepth ?? false,
          },
          title: Boolean(selectedChartDesignTitle),
          axisTitleX: Boolean(sheetChart.axisTitleX),
          axisTitleY: Boolean(sheetChart.axisTitleY),
          axisTitleDepth: sheetBaseView?.axisTitleDepth ?? false,
          axisX: sheetAxisShown("axisX"),
          axisY: sheetAxisShown("axisY"),
          // The model's own switch (the record's flat format has none).
          axisSecondaryY: sheetBaseView?.axisSecondaryY ?? false,
          // And the depth axis of a 3-D chart (`axisOptions.series.hidden`).
          axisDepth: sheetBaseView?.axisDepth ?? false,
          gridlinesX: sheetGridShown("gridlinesX"),
          gridlinesY: sheetGridShown("gridlinesY"),
          minorGridlinesX:
            sheetChartFormat?.layout?.minorGridlinesX === true ||
            sheetChartFormat?.minorGridlinesX === true,
          minorGridlinesY:
            sheetChartFormat?.layout?.minorGridlinesY === true ||
            sheetChartFormat?.minorGridlinesY === true,
          dataLabels: selectedChartDesignShowDataLabels,
          dataLabelPlacement: selectedDataLabelPlacement,
          dataLabelPlacements,
          dataLabelAutomatic: !dataLabelsAreMapLabels,
          dataTable: selectedChartDesignShowDataTable,
          dataTableLegendKeys: sheetChartFormat?.dataTableLegendKeys !== false,
          errorBars:
            sheetErrorAmount === "percent"
              ? "percentage"
              : sheetErrorAmount === "stddev"
                ? "standardDeviation"
                : sheetErrorAmount === "stderr"
                  ? "standardError"
                  : (sheetErrorAmount ?? null),
          legend: selectedChartDesignShowLegend
            ? (selectedChartDesignLegendPosition as OfficeChartLegendPosition)
            : "none",
          dropLines: sheetChart.dropLines === true,
          highLowLines: sheetChart.highLowLines === true,
          seriesLines: sheetChart.seriesLines === true,
          trendline: sheetAnalysis.trendline
            ? ({
                type: sheetAnalysis.trendline.type,
                ...(sheetAnalysis.trendline.type === "movingAverage"
                  ? { period: sheetAnalysis.trendline.period }
                  : {}),
                // Linear Forecast is the linear line forecast
                // forward two periods; the menu reads it back.
                ...(sheetAnalysis.trendline.forward
                  ? { forward: sheetAnalysis.trendline.forward }
                  : {}),
              } as ChartTrendline)
            : null,
          upDownBars: Boolean(sheetChart.upDownBars),
        };
        const runChartDesign = (
          controlId: Parameters<typeof executeChartDesignMutation>[0],
          parameter: string,
        ) => void executeChartDesignMutation(controlId, parameter);
        const errorBarsParameter = (
          type: "fixed" | "percent" | "stddev" | "stderr",
          preset: number,
        ) =>
          JSON.stringify({
            type,
            value:
              sheetAnalysis.errorBars?.type === type
                ? sheetAnalysis.errorBars.value
                : preset,
          });
        const onSheetChartElement = (intent: ChartElementIntent) => {
          switch (intent.element) {
            case "chartTitle":
              return runChartDesign(
                "XLSCTL.CONTEXT.CHART.TITLE.MRAI.APPLY",
                intent.on
                  ? selectedChartDesignTitle ||
                      t("spreadsheet.uiCopy.chartTitle")
                  : "",
              );
            case "axisTitle":
              if (intent.axis === "z") {
                return runChartDesign(
                  "XLSCTL.CONTEXT.CHART.AXIS_TITLE_DEPTH.MRAI.APPLY",
                  intent.on ? intent.text || t("office.chartElements.depthAxisTitle") : "",
                );
              }
              return runChartDesign(
                intent.axis === "x"
                  ? "XLSCTL.CONTEXT.CHART.AXIS_TITLE_X.MRAI.APPLY"
                  : "XLSCTL.CONTEXT.CHART.AXIS_TITLE_Y.MRAI.APPLY",
                intent.on
                  ? t(
                      intent.axis === "x"
                        ? "spreadsheet.uiCopy.horizontalAxisTitle"
                        : "spreadsheet.uiCopy.verticalAxisTitle",
                    )
                  : "",
              );
            case "axis":
              return runChartDesign(
                "XLSCTL.CONTEXT.CHART.AXES",
                `${intent.axis}:${intent.on}`,
              );
            case "gridlines":
              return runChartDesign(
                "XLSCTL.CONTEXT.CHART.GRIDLINES",
                `${intent.minor ? "minor-" : ""}${intent.axis}:${intent.on}`,
              );
            case "dataLabels":
              return intent.on
                ? runChartDesign(
                    "XLSCTL.CONTEXT.CHART.DATA_LABEL_POSITION.MRAI.APPLY",
                    intent.callout ? "callout" : (intent.placement ?? "automatic"),
                  )
                : runChartDesign("XLSCTL.CONTEXT.CHART.DATA_LABELS", "false");
            case "dataTable":
              return runChartDesign(
                "XLSCTL.CONTEXT.CHART.DATA_TABLE",
                !intent.on
                  ? "false"
                  : !intent.legendKeys
                    ? "true:no-keys"
                    : sheetChartFormat?.dataTableLegendKeys === false
                      ? "true:keys"
                      : "true",
              );
            case "errorBars":
              return runChartDesign(
                "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
                intent.recipe === null
                  ? "null"
                  : intent.recipe.type === "percentage"
                    ? errorBarsParameter("percent", intent.recipe.value ?? 5)
                    : intent.recipe.type === "standardDeviation"
                      ? errorBarsParameter("stddev", intent.recipe.value ?? 1)
                      : errorBarsParameter("stderr", 0),
              );
            case "legend":
              return intent.position === "none"
                ? runChartDesign("XLSCTL.CONTEXT.CHART.LEGEND", "false")
                : runChartDesign(
                    "XLSCTL.CONTEXT.CHART.LEGEND_POSITION.MRAI.APPLY",
                    intent.position,
                  );
            case "line": {
              const key =
                intent.line === "drop"
                  ? "dropLines"
                  : intent.line === "highLow"
                    ? "highLowLines"
                    : "seriesLines";
              const next: {
                dropLines?: true;
                highLowLines?: true;
                seriesLines?: true;
              } = {
                ...(sheetChart.dropLines ? { dropLines: true } : {}),
                ...(sheetChart.highLowLines ? { highLowLines: true } : {}),
                ...(sheetChart.seriesLines ? { seriesLines: true } : {}),
              };
              if (intent.on) next[key] = true;
              else delete next[key];
              return runChartDesign(
                "XLSCTL.CONTEXT.CHART.LINES.MRAI.APPLY",
                Object.keys(next).length ? JSON.stringify(next) : "null",
              );
            }
            case "linesNone":
              return runChartDesign("XLSCTL.CONTEXT.CHART.LINES.MRAI.APPLY", "null");
            case "trendline": {
              const type = intent.trendline?.type;
              // Linear Forecast carries its two forward periods
              // (`c:forward`); every other row states none.
              const forecast = intent.forward
                ? { forward: intent.forward }
                : {};
              return runChartDesign(
                "XLSCTL.CONTEXT.CHART.TRENDLINE.MRAI.APPLY",
                JSON.stringify(
                  !type
                    ? null
                    : type === "movingAverage"
                      ? {
                          type,
                          // Two Period Moving Average: two.
                          period:
                            (intent.trendline as { period?: number }).period ??
                            sheetAnalysis.trendline?.period ??
                            DEFAULT_TRENDLINE_PERIOD,
                        }
                      : type === "polynomial"
                        ? { type, order: sheetAnalysis.trendline?.order ?? 2 }
                        : { type, ...forecast },
                ),
              );
            }
            case "upDownBars":
              return runChartDesign(
                "XLSCTL.CONTEXT.CHART.UP_DOWN_BARS.MRAI.APPLY",
                String(intent.on),
              );
            case "dataLabelOverride":
              // The Format pane's, never the Add Chart Element menu's.
              return;
            case "displayBlanksAs":
              return runChartDesign(
                "XLSCTL.CONTEXT.CHART.DISPLAY_BLANKS.MRAI.APPLY",
                intent.value,
              );
            case "trendlineOption": {
              // Display Equation / R-squared and the forecast
              // ride on the trendline, through the same command.
              const next = desktopChartAttrs
                ? chartWithTrendlineOption(desktopChartAttrs, intent.change)
                : null;
              return next?.format?.trendline
                ? runChartDesign(
                    "XLSCTL.CONTEXT.CHART.TRENDLINE.MRAI.APPLY",
                    JSON.stringify(next.format.trendline),
                  )
                : undefined;
            }
          }
        };
        const sheetElementWords: Record<string, string> = {
          axes: t("spreadsheet.chartFormatPane.axes"),
          axisTitles: t("spreadsheet.uiCopy.axisTitles"),
          chartTitle: t("spreadsheet.uiCopy.chartTitle"),
          dataLabels: t("spreadsheet.uiCopy.dataLabels"),
          dataTable: t("spreadsheet.uiCopy.dataTable"),
          errorBars: t("spreadsheet.uiCopy.errorBars"),
          gridlines: t("toolbar.view.gridlines"),
          legend: t("spreadsheet.chartLegendPosition"),
          lines: t("presentation.shapeGallery.category.lines"),
          trendline: t("spreadsheet.chartTrendline"),
          upDownBars: t("spreadsheet.chartFormatPane.upDownBars"),
          none: t("spreadsheet.chartLineMarkerNone"),
          automatic: t("spreadsheet.chartFormatCopy.pane.automatic"),
          aboveChart: t("spreadsheet.uiCopy.aboveChart"),
          primaryHorizontal: t("spreadsheet.uiCopy.primaryHorizontal"),
          primaryVertical: t("spreadsheet.uiCopy.primaryVertical"),
          secondaryVertical: t("office.chartElements.secondaryVertical"),
          horizontalAxisTitle: t("spreadsheet.uiCopy.horizontalAxisTitle"),
          verticalAxisTitle: t("spreadsheet.uiCopy.verticalAxisTitle"),
          legendRight: t(SPREADSHEET_CHART_LEGEND_LABEL_KEYS.right),
          legendTop: t(SPREADSHEET_CHART_LEGEND_LABEL_KEYS.top),
          legendLeft: t(SPREADSHEET_CHART_LEGEND_LABEL_KEYS.left),
          legendBottom: t(SPREADSHEET_CHART_LEGEND_LABEL_KEYS.bottom),
          percentage: t("spreadsheet.chartErrorBarsPercent"),
          standardDeviation: t("spreadsheet.uiCopy.standardDeviation"),
          dropLines: t("spreadsheet.uiCopy.dropLines"),
          highLowLines: t("spreadsheet.uiCopy.highLowLines"),
          seriesLines: t("spreadsheet.uiCopy.seriesLines"),
          "trendline.linear": t("spreadsheet.chartTrendlineLinear"),
          "trendline.movingAverage": t("spreadsheet.uiCopy.movingAverage"),
        };
        // Add Chart Element's rows, Change Colors' palettes and
        // the Chart Styles gallery: the ribbon draws them, and so
        // do Excel's buttons beside the selected chart.
        const sheetElementItems = (
                    <OfficeChartElementMenuItems
                      locale={locale}
                      attrs={desktopChartAttrs}
                      view={sheetElementView}
                      writable
                      // The workbook's trendline states a forecast
                      // (`forward`), so Linear Forecast is live.
                      trendlineOptions
                      onIntent={onSheetChartElement}
                      entryState={sheetElementEntry}
                      label={(key) => sheetElementWords[key]}
                      testIdPrefix="spreadsheet-ribbon-chart"
                      testIds={spreadsheetChartDesignTestId}
                      moreOptions={Object.fromEntries(
                        (
                          [
                            ["axes", "axis:category"],
                            ["chartTitle", "chartTitleStyle"],
                            ["dataLabels", "dataLabelStyle"],
                            ["dataTable", "dataTableStyle"],
                            ["errorBars", "errorBarStyle"],
                            ["gridlines", "gridlines"],
                            ["legend", "legendStyle"],
                            ["trendline", "trendlineStyle"],
                            ["upDownBars", "upBarStyle"],
                          ] as const
                        ).map(([element, area]) => [
                          element,
                          () => {
                            chartFormatFocusTokenRef.current += 1;
                            setChartFormatFocus({
                              area,
                              token: chartFormatFocusTokenRef.current,
                            });
                            setRibbonTab("chartFormat");
                          },
                        ]),
                      )}
                      subContentClassName="docs-cupertino-menu w-max max-w-[240px]"
                      // Office's Add Chart Element is the same
                      // menu in the three editors: a Fixed Value
                      // error amount and a series' Secondary Axis
                      // are Format pane options (Format Error
                      // Bars > Error Amount, Format Data Series >
                      // Plot Series On), which the sheet's pane
                      // carries; the menu adds no rows of its own.
                    />
        );
        const sheetColorItems = (
                  <OfficeChartColorPaletteMenuItems
                    active={selectedChartColorPaletteId ?? null}
                    accents={workbookThemeAccentColors}
                    onPreview={(paletteId) =>
                      setChartColorPalettePreview(
                        paletteId && selectedChartDesignKey
                          ? { chartKey: selectedChartDesignKey, paletteId }
                          : null,
                      )
                    }
                    onSelect={(paletteId) =>
                      void executeChartDesignMutation(
                        "XLSCTL.CONTEXT.CHART.CHANGE_COLORS.MRAI.APPLY",
                        paletteId,
                      )
                    }
                    labels={{
                      group: (group) =>
                        group === "colorful"
                          ? t("spreadsheet.uiCopy.colorful")
                          : t("spreadsheet.uiCopy.monochromatic"),
                      palette: (group, index) =>
                        t(
                          group === "colorful"
                            ? "spreadsheet.dynamicCopy.chartPalette.colorful"
                            : "spreadsheet.dynamicCopy.chartPalette.monochromatic",
                          { index },
                        ),
                    }}
                    testIdPrefix="spreadsheet-chart-color"
                  />
        );
        const renderSheetChartStyleGallery = (slot?: OfficeChartStylesSlot) =>
          desktopChartAttrs && desktopChartPreviewAttrs ? <DocumentChartStyleGallery
                    application="spreadsheet" testIdPrefix={slot?.testIdPrefix ?? "spreadsheet-chart-style-gallery"}
                    attrs={desktopChartAttrs} previewAttrs={desktopChartPreviewAttrs}
                    // An imported chart's tiles draw from the same look the canvas uses,
                    // so a filled map shows its own outlines instead of a blank tile.
                    nativeLook={normalizeChartNativeLook(
                      (selectedChartDesignTarget?.chart as { nativeLook?: unknown } | undefined)?.nativeLook,
                    )}
                    // The tiles in the workbook's own theme colours, as the chart
                    // itself and Change Colors draw: they were drawn in Office's
                    // default blue and orange beside a chart in the theme's teal
                    // and orange, so none of them looked like the chart.
                    themeAccents={workbookThemeAccentColors}
                    chartSize={selectedChartStyleGallerySize}
                    labels={{
                      title: t("spreadsheet.sharedStyleCopy.chart.title"),
                      more: t("spreadsheet.sharedStyleCopy.chart.more"),
                      unavailable: t("spreadsheet.sharedStyleCopy.chart.unavailable"),
                    }}
                    labelFor={id=>documentChartStyleLabel(t,id)}
                    disabledReason={desktopChartModelReason??chartStyleDisabledReason??null}
                    onPreview={(id)=>{const onModel=Boolean(selectedChartModel);const format=id&&!onModel?desktopChartStyleFormat(id):null;setChartDesktopPreview(id&&(onModel||format)&&selectedChartDesignKey?{key:selectedChartDesignKey,source:selectedChartDesignTarget.chart,styleId:id,format}:null);}}
                    onCommit={(id)=>{if(!chartStyleCommandState.enabled)return;desktopChartStyleCommit(id);}}
                    {...(slot ? { menuTrigger: slot.menuTrigger, menuHeader: slot.menuHeader, menuBody: slot.menuBody } : {})}
                  /> : null;
        const addElementAuthority = visibleCommandDisclosureProps(
          "XLSCTL.CONTEXT.CHART.ADD_ELEMENT_MENU",
        );
        const selectDataAuthority = visibleCommandDisclosureProps(
          "XLSCTL.CONTEXT.CHART.SELECT_DATA",
        );
        sheetChartQuick.key = selectedChartDesignKey;
        sheetChartQuick.node = (
          <OfficeChartQuickButtons
            locale={locale}
            testIdPrefix="spreadsheet-chart-quick"
            disabledReason={
              desktopChartModelReason ??
              (addElementAuthority.disabled
                ? (addElementAuthority.disabledReason ??
                  (visibleCommandReason("spreadsheet.capability.commandUnavailable") ?? null))
                : null)
            }
            elements={sheetElementItems}
            renderStyles={renderSheetChartStyleGallery}
            colors={chartColorCommandState.enabled ? sheetColorItems : null}
            colorsUnavailableReason={chartColorDisabledReason}
            onClose={() => {
              setChartColorPalettePreview(null);
              setChartDesktopPreview(null);
            }}
            filters={{
              ...officeChartFilterNames(
                desktopChartAttrs?.data ?? { kind: "none" },
              ),
              // Hiding a series or a category is written as Excel
              // writes it (`c15:filtered…Series`, `c15:fullRef`).
              ...officeChartFilterControls(
                desktopChartAttrs,
                desktopChartModelReason ? null : patchDesktopChartFilters,
                desktopChartModelReason ?? undefined,
              ),
              onSelectData: selectDataAuthority.disabled
                ? undefined
                : () =>
                    void executeVisibleCommand(
                      "XLSCTL.CONTEXT.CHART.SELECT_DATA",
                    ),
            }}
            style={{ position: "absolute", top: 0, left: "calc(100% + 6px)", zIndex: 30 }}
          />
        );
        return [
          {
            // The one Chart Design tab Docs and the deck mount
            // (OfficeChartDesignRibbon), Excel's Location group
            // included; the sheet supplies its command authority.
            id: "chart-design-shared",
            label: "",
            unwrapped: true,
            content: (
              <OfficeChartDesignRibbon
                frame="groups"
                ribbon="chartDesign"
                locale={locale}
                testId="spreadsheet-ribbon-chart-design"
                testIdPrefix="spreadsheet-ribbon-chart"
                testIds={spreadsheetChartDesignTestId}
                controlSurface="chart-design"
                classes={{
                  group: spreadsheetHomeStyles.secondaryGroup,
                  // The same command row every other secondary tab
                  // draws (Data, Review, View): the horizontal
                  // stack, then the 80px band. Without the stack the
                  // class is a block, and Quick Layout and the style
                  // gallery fell under their neighbours.
                  commands: `note-editor-ribbon-hstack ${spreadsheetHomeStyles.secondaryGroupContent}`,
                }}
                labels={{
                  layoutGroup: t("spreadsheet.uiCopy.chartLayouts"),
                  stylesGroup: t("spreadsheet.uiCopy.chartStyles"),
                  dataGroup: t("spreadsheet.chartFormatPane.dataTab"),
                  // The group is "Type" in Excel (its one button is
                  // Change Chart Type); the shared ribbon's own word.
                  locationGroup: t("toolbar.notePosition"),
                  addChartElement: t("spreadsheet.uiCopy.addChartElement"),
                  changeColors: t("presentation.smartartContext.changeColors"),
                  switchRowColumn: t("presentation.chartSwitchRowColumn"),
                  selectData: t("presentation.chartSelectData"),
                  moveChart: t("spreadsheet.uiCopy.moveChart"),
                }}
                elements={{
                  ...visibleCommandDisclosureProps(
                    "XLSCTL.CONTEXT.CHART.ADD_ELEMENT_MENU",
                  ),
                  ...(desktopChartModelReason ? { disabled: true, disabledReason: desktopChartModelReason } : {}),
                  occurrenceId: "add-chart-element",
                  contentClassName: "docs-cupertino-menu w-[280px]",
                  items: sheetElementItems,
                }}
                quickLayout={{
                  control: (
              <div
                data-control-id="XLSCTL.CONTEXT.CHART.QUICK_LAYOUT.MRAI.APPLY"
                data-control-occurrence-id="quick-layout"
                data-control-surface="chart-design"
                data-testid="spreadsheet-ribbon-chart-quick-layout-authority"
                title={chartQuickLayoutDisabledReason}
              >
                <OfficeChartQuickLayoutGallery
                  application="spreadsheet"
                  calibrated={chartQuickLayoutCalibrated}
                  available={chartQuickLayoutRefusal}
                  disabledReason={
                    desktopChartModelReason ??
                    (chartQuickLayoutCommandState.enabled
                      ? null
                      : (chartQuickLayoutDisabledReason ?? (visibleCommandReason("spreadsheet.capability.commandUnavailable") ?? null)))
                  }
                  labels={{
                    ariaLabel: t("spreadsheet.uiCopy.chartQuickLayout"),
                    more: t("spreadsheet.uiCopy.moreLayouts"),
                    item: (index) =>
                      t("spreadsheet.dynamicCopy.chartLayout", { index }),
                  }}
                  className={
                    spreadsheetHomeStyles.contextualStyleGallery
                  }
                  menuMode="preview-only"
                  trigger={<OfficeHomeMenuTrigger variant="large" hasMenuIndicator icon={<OfficeSemanticIcon semanticId="chart.dataTable" size={24}/>}
                    label={t("spreadsheet.uiCopy.quickLayout")} disabled={!chartQuickLayoutCommandState.enabled} disabledReason={chartQuickLayoutDisabledReason}
                    testId="spreadsheet-chart-quick-layout-trigger"/>}
                  testIdPrefix="spreadsheet-chart-quick-layout-gallery"
                  value={selectedChartQuickLayoutId}
                  onPreview={(layoutId) =>
                    setChartQuickLayoutPreview(
                      layoutId && selectedChartDesignKey
                        ? {
                            chartKey: selectedChartDesignKey,
                            layoutId,
                          }
                        : null,
                    )
                  }
                  onCommit={(layoutId) => {
                    setChartQuickLayoutPreview(null);
                    void executeChartDesignMutation(
                      "XLSCTL.CONTEXT.CHART.QUICK_LAYOUT.MRAI.APPLY",
                      layoutId,
                    );
                  }}
                />
              </div>
                  ),
                }}
                colors={{
                  controlId: "XLSCTL.CONTEXT.CHART.CHANGE_COLORS.MRAI.APPLY",
                  occurrenceId: "change-colors",
                  active: selectedChartColorPaletteId !== undefined,
                  disabled: !chartColorCommandState.enabled || Boolean(desktopChartModelReason),
                  disabledReason: desktopChartModelReason ?? chartColorDisabledReason,
                  contentClassName: "docs-cupertino-menu w-[310px]",
                  onCloseAutoFocus: () => setChartColorPalettePreview(null),
                  items: sheetColorItems,
                }}
                styleGallery={
                <div
                  data-control-id="XLSCTL.CONTEXT.CHART.STYLE_GALLERY.MRAI.APPLY"
                  data-control-occurrence-id="style-gallery"
                  data-control-surface="chart-design"
                  data-testid="spreadsheet-ribbon-chart-style-authority"
                  title={chartStyleDisabledReason}
                >
                  {renderSheetChartStyleGallery()}
                </div>
                }
                hiddenEmptyCells={
                  desktopChartAttrs
                    ? officeChartHiddenEmptyCellsMenu(desktopChartAttrs, {
                        translate: t,
                        ...chartDesignActionProps(
                          "XLSCTL.CONTEXT.CHART.DISPLAY_BLANKS.MRAI.APPLY",
                        ),
                        occurrenceId: "hidden-empty-cells",
                        onIntent: (intent) => {
                          if (intent.element === "displayBlanksAs")
                            void executeChartDesignMutation(
                              "XLSCTL.CONTEXT.CHART.DISPLAY_BLANKS.MRAI.APPLY",
                              intent.value,
                            );
                        },
                      })
                    : undefined
                }
                switchRowColumn={{
                  ...chartDesignActionProps(
                    "XLSCTL.CONTEXT.CHART.SWITCH_ROW_COLUMN",
                  ),
                  occurrenceId: "switch-row-column",
                  onClick: () =>
                    void executeChartDesignMutation(
                      "XLSCTL.CONTEXT.CHART.SWITCH_ROW_COLUMN",
                      String(sheetChartPlotsByRows(selectedChartDesignTarget)),
                    ),
                }}
                selectData={{
                  ...chartDesignActionProps(
                    "XLSCTL.CONTEXT.CHART.SELECT_DATA",
                  ),
                  occurrenceId: "select-data",
                  buttonRef: chartDataReturnFocusRef,
                  onClick: () =>
                    void executeVisibleCommand(
                      "XLSCTL.CONTEXT.CHART.SELECT_DATA",
                    ),
                }}
                changeType={{
                  control: <SpreadsheetChartTypeMenu translate={t} disabledReason={desktopChartModelReason}
              current={{type:chartType,subtype:desktopChartAttrs?.chartSubtype}} choice={desktopTypeChoice}
              label={(family,subtype)=>t(`toolbar.chartV2.subtype.${family}.${subtype}` as "toolbar.chartV2.subtype.column.clustered")}
              onChoose={(choice)=>void chooseDesktopChartType(choice)} />,
                }}
                location={{
                  buttonRef: chartMoveTriggerRef,
                  disabled:
                    Boolean(desktopChartFormatReason) ||
                    workbookTopologyLocked,
                  onClick: openDesktopChartMove,
                }}
              />
            ),
          },
        ];
      })()
    : [];
  return {
    chartDesignRibbonGroups,
  };
}
