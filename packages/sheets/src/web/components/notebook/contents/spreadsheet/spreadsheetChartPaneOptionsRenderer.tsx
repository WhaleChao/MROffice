"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { SpreadsheetSeriesIndicesInput } from "./SpreadsheetSeriesIndicesInput";
import { sheetChartPlotsByRows } from "../../../../lib/spreadsheet/drawing-buttons";
import { type ReactNode, Children } from "react";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import spreadsheetHomeStyles from "./SpreadsheetHomeRibbon.module.css";
import { type DocumentChartAreaStyleKey } from "@mrnote/core/web/components/console/editor/native-document/DocumentChartFormatRibbon.tsx";
import { type SpreadsheetVisibleCommandReason } from "../../../../lib/spreadsheet/visible-command-authority";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { applyNumberFormat } from "@mrnote/core/web/lib/spreadsheet/number-format.ts";
import { normalizeChartValueAxisNumberFormat, normalizeTrendline, type SpreadsheetLineMarkerSymbol, SPREADSHEET_LINE_MARKER_SYMBOLS, MAX_CHART_TITLE_LENGTH, DEFAULT_TRENDLINE_PERIOD, type ChartErrorBarType, DEFAULT_ERROR_BAR_VALUE, type SpreadsheetChart } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { type SpreadsheetObject } from "./spreadsheetSheetModel";
import { type ChartDesignMutationReceipt } from "./spreadsheetMutationReceipts";
import { type SpreadsheetChartLineSeriesStyle, type SpreadsheetChartLinkedFormatResolution, type SpreadsheetChartValueAxisNumberFormat } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { type Dispatch, type SetStateAction } from "react";
import { type ChartBlockAttrsV2 } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { chartElementAvailability, chartSeriesTrendline } from "@mrnote/core/web/lib/office-charts/chart-elements.ts";
import { chartSeriesPlottedOn } from "@mrnote/core/web/lib/office-charts/chart-type.ts";

/** What spreadsheetChartPaneOptionsRenderer reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetChartPaneOptionsRendererInputs {
  chartDesignControlState: (controlId: string) => { state: { enabled: boolean; visible: boolean; reasonKey: SpreadsheetVisibleCommandReason | null; featureId: string; controlId: string; }; reason: string | undefined; dom: { "data-control-id": string; "data-control-surface": string; disabled: boolean; hidden: boolean; title: string | undefined; }; };
  chartMarkerSeriesIndex: number;
  /** The selected chart's model, which states each series' own trendline. */
  desktopChartAttrs: ChartBlockAttrsV2 | null;
  executeChartDesignMutation: (controlId: ChartDesignMutationReceipt["controlId"], parameter: string) => Promise<boolean>;
  executeVisibleCommand: (controlId: string) => Promise<boolean>;
  selectedChartDesignKey: string | null;
  selectedChartDesignRange: string;
  selectedChartDesignTarget: { readonly kind: "drawing"; readonly chart: SpreadsheetObject; } | { readonly kind: "native"; readonly chart: SpreadsheetChart; } | null;
  selectedChartDesignTitle: string;
  selectedChartEffectiveValueAxisNumberFormat: string;
  selectedChartLineSeriesStyles: readonly SpreadsheetChartLineSeriesStyle[];
  selectedChartLinkedNumberFormatReason: string | null;
  selectedChartLinkedNumberFormatResolution: SpreadsheetChartLinkedFormatResolution;
  selectedChartMarkerSeriesIndices: number[];
  selectedChartValueAxisNumberFormat: SpreadsheetChartValueAxisNumberFormat;
  selectedChartValueAxisNumberFormatPaneVisible: boolean;
  setChartMarkerSeriesIndex: Dispatch<SetStateAction<number>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  visibleCommandDomStateProps: (controlId: string) => { "data-control-id": string; disabled: boolean; hidden: boolean; title: string | undefined; };
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function spreadsheetChartPaneOptionsRenderer({
  chartDesignControlState,
  chartMarkerSeriesIndex,
  desktopChartAttrs,
  executeChartDesignMutation,
  executeVisibleCommand,
  selectedChartDesignKey,
  selectedChartDesignRange,
  selectedChartDesignTarget,
  selectedChartDesignTitle,
  selectedChartEffectiveValueAxisNumberFormat,
  selectedChartLineSeriesStyles,
  selectedChartLinkedNumberFormatReason,
  selectedChartLinkedNumberFormatResolution,
  selectedChartMarkerSeriesIndices,
  selectedChartValueAxisNumberFormat,
  selectedChartValueAxisNumberFormatPaneVisible,
  setChartMarkerSeriesIndex,
  t,
  visibleCommandDomStateProps,
}: SpreadsheetChartPaneOptionsRendererInputs) {
  // The worksheet's own options on the shared Format pane, Office's
  // per-element options page ("Format Axis > Axis Options > Number", "Format
  // Data Series > Series Options", the data behind the chart area). They were
  // a second, older Series Options pane; now the Format pane shows them for
  // the element it is formatting.
  const renderSheetChartPaneOptions = (
    area: DocumentChartAreaStyleKey,
  ): ReactNode => {
    if (!selectedChartDesignTarget) return null;
    const seriesArea = area.startsWith("series:");
    // Format Data Series formats ONE series: its trendline is that series'
    // (Office bases a trendline on the series picked), never the first's.
    const paneSeriesIndex = seriesArea ? Number(area.slice("series:".length)) : null;
    const paneTrendline: { type: string; period?: number } | null | undefined =
      paneSeriesIndex !== null && desktopChartAttrs
        ? chartSeriesTrendline(desktopChartAttrs, paneSeriesIndex)
        : normalizeTrendline(selectedChartDesignTarget.chart.trendline);
    const writePaneTrendline = (trendline: object | null) =>
      void executeChartDesignMutation(
        "XLSCTL.CONTEXT.CHART.TRENDLINE.MRAI.APPLY",
        JSON.stringify(
          paneSeriesIndex !== null
            ? { seriesIndex: paneSeriesIndex, trendline }
            : trendline,
        ),
      );
    // The chart's error bars are its first series'; another series' pane
    // offers none rather than formatting the first series in its name.
    // The shared Format pane's Trendline and Error Bars pages (and a series'
    // own trendline) write these wherever the chart's model takes them; the
    // worksheet's own controls stay only for its overlay on a family the
    // model refuses them on (a combo, a stacked kind).
    const sharedElements = desktopChartAttrs ? chartElementAvailability(desktopChartAttrs) : null;
    const paneOffersTrendline = !sharedElements?.trendline;
    const paneOffersErrorBars =
      (paneSeriesIndex === null || paneSeriesIndex === 0) && !sharedElements?.errorBars;
    // Only the element's own options: an element the worksheet has none for
    // returns nothing, so the shared pane shows no empty Options page.
    const options = (
      <>
          {area === "axis:value" &&
          selectedChartValueAxisNumberFormatPaneVisible
            ? (() => {
                const authority = chartDesignControlState(
                  "XLSCTL.CONTEXT.CHART.VALUE_AXIS_NUMBER_FORMAT.MRAI.APPLY",
                );
                const linked =
                  selectedChartValueAxisNumberFormat.sourceLinked;
                const paneReason = !authority.state.enabled
                  ? authority.reason
                  : linked
                    ? selectedChartLinkedNumberFormatReason
                    : null;
                const preview = applyNumberFormat(
                  1234.5,
                  selectedChartEffectiveValueAxisNumberFormat,
                );
                return (
                  <section
                    className={spreadsheetHomeStyles.chartFormatSection}
                    data-testid="spreadsheet-chart-value-axis-number-format"
                  >
                    <h3
                      className={
                        spreadsheetHomeStyles.chartFormatSectionHeading
                      }
                    >
                      {t("spreadsheet.chartValueAxisNumberFormat")}
                    </h3>
                    <div className={spreadsheetHomeStyles.chartFormatFields}>
                      <label className={spreadsheetHomeStyles.chartFormatField}>
                        <span>
                          {t("spreadsheet.chartValueAxisNumberFormatSource")}
                        </span>
                        <select
                          {...authority.dom}
                          aria-label={t(
                            "spreadsheet.chartValueAxisNumberFormatSource",
                          )}
                          data-control-occurrence-id="format-value-axis-number-format-source"
                          data-control-surface="chart-format-pane"
                          data-testid="spreadsheet-chart-value-axis-number-format-source"
                          value={linked ? "linked" : "custom"}
                          onChange={(event) => {
                            const sourceLinked =
                              event.currentTarget.value === "linked";
                            void executeChartDesignMutation(
                              "XLSCTL.CONTEXT.CHART.VALUE_AXIS_NUMBER_FORMAT.MRAI.APPLY",
                              JSON.stringify({
                                formatCode: sourceLinked
                                  ? selectedChartLinkedNumberFormatResolution.formatCode
                                  : selectedChartEffectiveValueAxisNumberFormat,
                                sourceLinked,
                              }),
                            );
                          }}
                        >
                          <option
                            value="linked"
                            disabled={
                              selectedChartLinkedNumberFormatResolution.state !==
                              "resolved"
                            }
                          >
                            {t(
                              "spreadsheet.chartValueAxisNumberFormatLinked",
                            )}
                          </option>
                          <option value="custom">
                            {t(
                              "spreadsheet.chartValueAxisNumberFormatCustom",
                            )}
                          </option>
                        </select>
                      </label>
                      <label className={spreadsheetHomeStyles.chartFormatField}>
                        <span>{t("spreadsheet.chartValueAxisFormatCode")}</span>
                        <input
                          key={`${selectedChartDesignKey}:${linked}:${selectedChartEffectiveValueAxisNumberFormat}`}
                          {...authority.dom}
                          aria-describedby={
                            paneReason
                              ? "spreadsheet-chart-value-axis-number-format-reason"
                              : undefined
                          }
                          aria-label={t("spreadsheet.chartValueAxisFormatCode")}
                          data-control-occurrence-id="format-value-axis-number-format-code"
                          data-control-surface="chart-format-pane"
                          data-testid="spreadsheet-chart-value-axis-number-format-code"
                          defaultValue={
                            selectedChartEffectiveValueAxisNumberFormat
                          }
                          disabled={authority.dom.disabled || linked}
                          maxLength={255}
                          title={
                            linked
                              ? paneReason ??
                                t(
                                  "spreadsheet.chartValueAxisNumberFormatLinkedHint",
                                )
                              : authority.reason
                          }
                          onBlur={(event) => {
                            const next = normalizeChartValueAxisNumberFormat({
                              formatCode: event.currentTarget.value,
                              sourceLinked: false,
                            });
                            if (!next) {
                              event.currentTarget.value =
                                selectedChartValueAxisNumberFormat.formatCode;
                              return;
                            }
                            void executeChartDesignMutation(
                              "XLSCTL.CONTEXT.CHART.VALUE_AXIS_NUMBER_FORMAT.MRAI.APPLY",
                              JSON.stringify(next),
                            );
                          }}
                          onKeyDown={(event) => {
                            if (event.key === "Enter") {
                              event.preventDefault();
                              event.currentTarget.blur();
                            } else if (event.key === "Escape") {
                              event.preventDefault();
                              event.currentTarget.value =
                                selectedChartEffectiveValueAxisNumberFormat;
                              event.currentTarget.blur();
                            }
                          }}
                        />
                      </label>
                      <output
                        aria-live="polite"
                        data-format-code={
                          selectedChartEffectiveValueAxisNumberFormat
                        }
                        data-format-state={
                          linked
                            ? selectedChartLinkedNumberFormatResolution.state
                            : "custom"
                        }
                        data-testid="spreadsheet-chart-value-axis-number-format-preview"
                        style={{ fontSize: 12 }}
                      >
                        {t("spreadsheet.chartValueAxisNumberFormatPreview", {
                          value: preview,
                        })}
                      </output>
                      {paneReason ? (
                        <p
                          id="spreadsheet-chart-value-axis-number-format-reason"
                          data-testid="spreadsheet-chart-value-axis-number-format-reason"
                          style={{
                            color: "var(--muted-foreground, #64748b)",
                            fontSize: 12,
                            margin: 0,
                          }}
                        >
                          {paneReason}
                        </p>
                      ) : null}
                    </div>
                  </section>
                );
              })()
            : null}
                {/* The shared Fill & Line page formats a series' marker
                    wherever the chart has its model; this list stays for a
                    chart that has none. */}
                {seriesArea &&
                !desktopChartAttrs &&
                selectedChartMarkerSeriesIndices.length > 0
                  ? (() => {
                      const symbolAuthority = chartDesignControlState(
                        "XLSCTL.CONTEXT.CHART.LINE_MARKER_SYMBOL.MRAI.APPLY",
                      );
                      const sizeAuthority = chartDesignControlState(
                        "XLSCTL.CONTEXT.CHART.LINE_MARKER_SIZE.MRAI.APPLY",
                      );
                      // A combo's markers belong to its line plot, so the pane
                      // opens on a series that has them, never on its columns.
                      const markerSeriesIndex =
                        selectedChartMarkerSeriesIndices.includes(
                          chartMarkerSeriesIndex,
                        )
                          ? chartMarkerSeriesIndex
                          : selectedChartMarkerSeriesIndices[0];
                      const style = selectedChartLineSeriesStyles.find(
                        (entry) => entry.seriesIndex === markerSeriesIndex,
                      );
                      const markerSymbol = style?.markerSymbol ?? "circle";
                      const markerSize = style?.markerSize ?? 6;
                      const markerLabels: Record<
                        SpreadsheetLineMarkerSymbol,
                        string
                      > = {
                        none: t("spreadsheet.chartLineMarkerNone"),
                        circle: t("spreadsheet.chartLineMarkerCircle"),
                        dash: t("spreadsheet.chartLineMarkerDash"),
                        diamond: t("spreadsheet.chartLineMarkerDiamond"),
                        dot: t("spreadsheet.chartLineMarkerDot"),
                        plus: t("spreadsheet.chartLineMarkerPlus"),
                        square: t("spreadsheet.chartLineMarkerSquare"),
                        star: t("spreadsheet.chartLineMarkerStar"),
                        triangle: t("spreadsheet.chartLineMarkerTriangle"),
                        x: "X",
                      };
                      return (
                        <section
                          className={spreadsheetHomeStyles.chartFormatSection}
                          data-testid="spreadsheet-chart-line-marker"
                        >
                          <h3
                            className={
                              spreadsheetHomeStyles.chartFormatSectionHeading
                            }
                          >
                            {t("spreadsheet.chartLineMarker")}
                          </h3>
                          <div
                            className={spreadsheetHomeStyles.chartFormatFields}
                          >
                            <label
                              className={spreadsheetHomeStyles.chartFormatField}
                            >
                              <span>
                                {t("spreadsheet.chartLineMarkerSeries")}
                              </span>
                              <select
                                aria-label={t("spreadsheet.chartLineMarkerSeries")}
                                data-testid="spreadsheet-chart-line-marker-series"
                                value={markerSeriesIndex}
                                onChange={(event) =>
                                  setChartMarkerSeriesIndex(
                                    Number.parseInt(
                                      event.currentTarget.value,
                                      10,
                                    ),
                                  )
                                }
                              >
                                {selectedChartMarkerSeriesIndices.map(
                                  (index) => (
                                    <option key={index} value={index}>
                                      {t("spreadsheet.chartLineMarkerSeriesNumber", {
                                        number: index + 1,
                                      })}
                                    </option>
                                  ),
                                )}
                              </select>
                            </label>
                            <label
                              className={spreadsheetHomeStyles.chartFormatField}
                            >
                              <span>
                                {t("spreadsheet.chartLineMarkerSymbol")}
                              </span>
                              <select
                                {...symbolAuthority.dom}
                                aria-label={t("spreadsheet.chartLineMarkerSymbol")}
                                data-control-occurrence-id="format-line-marker-symbol"
                                data-control-surface="chart-format-pane"
                                data-testid="spreadsheet-chart-line-marker-symbol"
                                value={markerSymbol}
                                onChange={(event) =>
                                  void executeChartDesignMutation(
                                    "XLSCTL.CONTEXT.CHART.LINE_MARKER_SYMBOL.MRAI.APPLY",
                                    JSON.stringify({
                                      seriesIndex: markerSeriesIndex,
                                      markerSymbol: event.currentTarget.value,
                                    }),
                                  )
                                }
                              >
                                {SPREADSHEET_LINE_MARKER_SYMBOLS.map(
                                  (symbol) => (
                                    <option key={symbol} value={symbol}>
                                      {markerLabels[symbol]}
                                    </option>
                                  ),
                                )}
                              </select>
                            </label>
                            <label
                              className={spreadsheetHomeStyles.chartFormatField}
                            >
                              <span>
                                {t("spreadsheet.chartLineMarkerSize")}
                              </span>
                              <input
                                {...sizeAuthority.dom}
                                aria-label={t("spreadsheet.chartLineMarkerSize")}
                                data-control-occurrence-id="format-line-marker-size"
                                data-control-surface="chart-format-pane"
                                data-testid="spreadsheet-chart-line-marker-size"
                                disabled={
                                  sizeAuthority.dom.disabled ||
                                  markerSymbol === "none"
                                }
                                max={72}
                                min={2}
                                step={1}
                                type="number"
                                value={markerSize}
                                onChange={(event) => {
                                  const next = Number.parseInt(
                                    event.currentTarget.value,
                                    10,
                                  );
                                  if (!Number.isInteger(next)) return;
                                  void executeChartDesignMutation(
                                    "XLSCTL.CONTEXT.CHART.LINE_MARKER_SIZE.MRAI.APPLY",
                                    JSON.stringify({
                                      seriesIndex: markerSeriesIndex,
                                      markerSize: next,
                                    }),
                                  );
                                }}
                              />
                            </label>
                          </div>
                        </section>
                      );
                    })()
                  : null}
          {area === "chartAreaStyle" || area === "chartTitleStyle" ? (
            <section className={spreadsheetHomeStyles.chartFormatSection}>
                    <h3
                      className={
                        spreadsheetHomeStyles.chartFormatSectionHeading
                      }
                    >
                {t("spreadsheet.chartFormatPane.titlesText")}
              </h3>
              <div className={spreadsheetHomeStyles.chartFormatFields}>
                {(
                  [
                    [
                      "XLSCTL.CONTEXT.CHART.TITLE.MRAI.APPLY",
                      t("toolbar.chartSampleTitle"),
                      selectedChartDesignTitle,
                      "format-title",
                    ],
                    [
                      "XLSCTL.CONTEXT.CHART.AXIS_TITLE_X.MRAI.APPLY",
                      t("presentation.chartAxisTitleX"),
                      selectedChartDesignTarget.chart.axisTitleX ?? "",
                      "format-axis-title-x",
                    ],
                    [
                      "XLSCTL.CONTEXT.CHART.AXIS_TITLE_Y.MRAI.APPLY",
                      t("presentation.chartAxisTitleY"),
                      selectedChartDesignTarget.chart.axisTitleY ?? "",
                      "format-axis-title-y",
                    ],
                  ] as const
                ).map(([controlId, label, value, occurrenceId]) => {
                  const authority = chartDesignControlState(controlId);
                  return (
                    <label
                      key={controlId}
                      className={spreadsheetHomeStyles.chartFormatField}
                    >
                      <span>{label}</span>
                      <input
                        key={`${selectedChartDesignKey}:${controlId}:${value}`}
                        {...authority.dom}
                        aria-label={
                          authority.reason
                            ? `${label}: ${authority.reason}`
                            : label
                        }
                        data-control-occurrence-id={occurrenceId}
                        defaultValue={value}
                        maxLength={MAX_CHART_TITLE_LENGTH}
                        onBlur={(event) =>
                          void executeChartDesignMutation(
                            controlId,
                            event.currentTarget.value,
                          )
                        }
                        onKeyDown={(event) => {
                          if (event.key === "Enter") {
                            event.preventDefault();
                            event.currentTarget.blur();
                          } else if (event.key === "Escape") {
                            event.preventDefault();
                            event.currentTarget.value = value;
                            event.currentTarget.blur();
                          }
                        }}
                      />
                    </label>
                  );
                })}
              </div>
            </section>
          ) : null}
          {(seriesArea ||
            area === "trendlineStyle" ||
            area === "errorBarStyle" ||
            area === "upBarStyle" ||
            area === "downBarStyle") &&
          selectedChartDesignTarget.kind === "native" ? (
            <section className={spreadsheetHomeStyles.chartFormatSection}>
              <h3 className={spreadsheetHomeStyles.chartFormatSectionHeading}>
                {t("spreadsheet.chartFormatPane.seriesOptions")}
              </h3>
              <div className={spreadsheetHomeStyles.chartFormatFields}>
              {paneOffersTrendline ? (<>
              <label className="grid gap-1.5 font-medium">
                <span>{t("spreadsheet.chartTrendline")}</span>
                <select
                  {...chartDesignControlState(
                    "XLSCTL.CONTEXT.CHART.TRENDLINE.MRAI.APPLY",
                  ).dom}
                  data-testid={`spreadsheet-chart-trendline-${selectedChartDesignTarget.chart.id}`}
                  value={paneTrendline?.type ?? "none"}
                  onChange={(event) => {
                    const value = event.currentTarget.value;
                    writePaneTrendline(
                      value === "linear"
                        ? { type: "linear" }
                        : value === "movingAverage"
                          ? {
                              type: "movingAverage",
                              period:
                                paneTrendline?.period ?? DEFAULT_TRENDLINE_PERIOD,
                            }
                          : null,
                    );
                  }}
                >
                  <option value="none">
                    {t("spreadsheet.chartTrendlineNone")}
                  </option>
                  <option value="linear">
                    {t("spreadsheet.chartTrendlineLinear")}
                  </option>
                  <option value="movingAverage">
                    {t("spreadsheet.chartTrendlineMovingAverage")}
                  </option>
                </select>
              </label>
              {paneTrendline?.type === "movingAverage" ? (
                <label className="grid gap-1.5 font-medium">
                  <span>{t("spreadsheet.chartTrendlinePeriod")}</span>
                  <input
                    {...chartDesignControlState(
                      "XLSCTL.CONTEXT.CHART.TRENDLINE.MRAI.APPLY",
                    ).dom}
                    data-testid={`spreadsheet-chart-trendline-period-${selectedChartDesignTarget.chart.id}`}
                    min={2}
                    step={1}
                    type="number"
                    value={
                      paneTrendline.period ?? DEFAULT_TRENDLINE_PERIOD
                    }
                    onChange={(event) => {
                      const parsed = Number.parseInt(
                        event.currentTarget.value,
                        10,
                      );
                      writePaneTrendline({
                        type: "movingAverage",
                        period: Number.isFinite(parsed)
                          ? Math.max(DEFAULT_TRENDLINE_PERIOD, parsed)
                          : DEFAULT_TRENDLINE_PERIOD,
                      });
                    }}
                  />
                </label>
              ) : null}
              </>) : null}
              {paneOffersErrorBars ? (<>
              <label className="grid gap-1.5 font-medium">
                <span>{t("spreadsheet.chartErrorBars")}</span>
                <select
                  {...chartDesignControlState(
                    "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
                  ).dom}
                  data-testid={`spreadsheet-chart-error-bars-${selectedChartDesignTarget.chart.id}`}
                  value={
                    selectedChartDesignTarget.chart.errorBars?.type ?? "none"
                  }
                  onChange={(event) => {
                    const value = event.currentTarget.value;
                    void executeChartDesignMutation(
                      "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
                      JSON.stringify(
                        value === "none"
                          ? null
                          : {
                              ...selectedChartDesignTarget.chart.errorBars,
                              type: value as ChartErrorBarType,
                              value:
                                selectedChartDesignTarget.chart.errorBars
                                  ?.value ?? DEFAULT_ERROR_BAR_VALUE,
                            },
                      ),
                    );
                  }}
                >
                  <option value="none">
                    {t("spreadsheet.chartErrorBarsNone")}
                  </option>
                  <option value="fixed">
                    {t("spreadsheet.chartErrorBarsFixed")}
                  </option>
                  <option value="percent">
                    {t("spreadsheet.chartErrorBarsPercent")}
                  </option>
                  <option value="stddev">
                    {t("spreadsheet.chartErrorBarsStdDev")}
                  </option>
                  <option value="stderr">
                    {t("spreadsheet.chartErrorBarsStdError")}
                  </option>
                </select>
              </label>
              {selectedChartDesignTarget.chart.errorBars &&
              selectedChartDesignTarget.chart.errorBars.type !== "stderr" ? (
                <label className="grid gap-1.5 font-medium">
                  <span>{t("spreadsheet.chartErrorBarsValue")}</span>
                  <input
                    {...chartDesignControlState(
                      "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
                    ).dom}
                    data-testid={`spreadsheet-chart-error-bars-value-${selectedChartDesignTarget.chart.id}`}
                    min={0}
                    step="any"
                    type="number"
                    value={selectedChartDesignTarget.chart.errorBars.value}
                    onChange={(event) => {
                      const parsed = Number.parseFloat(
                        event.currentTarget.value,
                      );
                      void executeChartDesignMutation(
                        "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
                        JSON.stringify({
                          ...selectedChartDesignTarget.chart.errorBars,
                          type: selectedChartDesignTarget.chart.errorBars!.type,
                          value: Number.isFinite(parsed)
                            ? Math.max(0, parsed)
                            : 0,
                        }),
                      );
                    }}
                  />
                </label>
              ) : null}
              {selectedChartDesignTarget.chart.errorBars ? (
                <>
                  {/* Format Error Bars > Direction and End Style. */}
                  <label className="grid gap-1.5 font-medium">
                    <span>{t("spreadsheet.chartErrorBarsDirection")}</span>
                    <select
                      {...chartDesignControlState(
                        "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
                      ).dom}
                      data-testid={`spreadsheet-chart-error-bars-direction-${selectedChartDesignTarget.chart.id}`}
                      value={selectedChartDesignTarget.chart.errorBars.include ?? "both"}
                      onChange={(event) => {
                        const include = event.currentTarget.value;
                        const { include: _include, ...rest } =
                          selectedChartDesignTarget.chart.errorBars!;
                        void executeChartDesignMutation(
                          "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
                          JSON.stringify(
                            include === "plus" || include === "minus"
                              ? { ...rest, include }
                              : rest,
                          ),
                        );
                      }}
                    >
                      <option value="both">{t("spreadsheet.chartErrorBarsBoth")}</option>
                      <option value="minus">{t("spreadsheet.chartErrorBarsMinus")}</option>
                      <option value="plus">{t("spreadsheet.chartErrorBarsPlus")}</option>
                    </select>
                  </label>
                  <label className="grid gap-1.5 font-medium">
                    <span>{t("spreadsheet.chartErrorBarsEndStyle")}</span>
                    <select
                      {...chartDesignControlState(
                        "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
                      ).dom}
                      data-testid={`spreadsheet-chart-error-bars-end-style-${selectedChartDesignTarget.chart.id}`}
                      value={selectedChartDesignTarget.chart.errorBars.noEndCap ? "none" : "cap"}
                      onChange={(event) => {
                        const { noEndCap: _noEndCap, ...rest } =
                          selectedChartDesignTarget.chart.errorBars!;
                        void executeChartDesignMutation(
                          "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
                          JSON.stringify(
                            event.currentTarget.value === "none"
                              ? { ...rest, noEndCap: true }
                              : rest,
                          ),
                        );
                      }}
                    >
                      <option value="none">{t("spreadsheet.chartErrorBarsNoCap")}</option>
                      <option value="cap">{t("spreadsheet.chartErrorBarsCap")}</option>
                    </select>
                  </label>
                </>
              ) : null}
              </> ) : null}
              <label className="grid gap-1.5 font-medium">
                <span>{t("spreadsheet.chartFormatPane.upDownBars")}</span>
                <select
                  {...chartDesignControlState(
                    "XLSCTL.CONTEXT.CHART.UP_DOWN_BARS.MRAI.APPLY",
                  ).dom}
                  data-testid={`spreadsheet-chart-up-down-bars-${selectedChartDesignTarget.chart.id}`}
                  value={
                    selectedChartDesignTarget.chart.upDownBars
                      ? "shown"
                      : "none"
                  }
                  onChange={(event) =>
                    void executeChartDesignMutation(
                      "XLSCTL.CONTEXT.CHART.UP_DOWN_BARS.MRAI.APPLY",
                      String(event.currentTarget.value === "shown"),
                    )
                  }
                >
                  <option value="none">
                    {t("spreadsheet.chartLineMarkerNone")}
                  </option>
                  <option value="shown">
                    {t("spreadsheet.chartFormatPane.upDownBars")}
                  </option>
                </select>
              </label>
              {/* Format Data Series ▸ Plot Series On is the shared page's
                  wherever the chart offers it. */}
              {desktopChartAttrs && chartSeriesPlottedOn(desktopChartAttrs, 0) !== null ? null : (
              <label className="grid gap-1.5 font-medium">
                <span>{t("spreadsheet.chartSecondaryAxis")}</span>
                <SpreadsheetSeriesIndicesInput
                  key={selectedChartDesignKey}
                  {...chartDesignControlState(
                    "XLSCTL.CONTEXT.CHART.SECONDARY_AXIS.MRAI.APPLY",
                  ).dom}
                  data-testid={`spreadsheet-chart-secondary-axis-${selectedChartDesignTarget.chart.id}`}
                  inputMode="numeric"
                  placeholder={t("spreadsheet.chartSecondaryAxisPlaceholder")}
                  indices={selectedChartDesignTarget.chart.secondaryAxisSeries ?? []}
                  onCommit={(indices) => void executeChartDesignMutation(
                    "XLSCTL.CONTEXT.CHART.SECONDARY_AXIS.MRAI.APPLY", JSON.stringify(indices),
                  )}
                />
              </label>
              )}
              </div>
            </section>
          ) : null}
          {area === "chartAreaStyle" ? (
            <section className={spreadsheetHomeStyles.chartFormatSection}>
              <h3 className={spreadsheetHomeStyles.chartFormatSectionHeading}>
                {t("spreadsheet.chartFormatPane.dataSource")}
              </h3>
              <div className={spreadsheetHomeStyles.chartFormatFields}>
                <label className={spreadsheetHomeStyles.chartFormatField}>
                  <span>{t("spreadsheet.chartDataRange")}</span>
                  <output
                    className={spreadsheetHomeStyles.chartFormatDataRange}
                    data-testid="spreadsheet-chart-format-range"
                    title={selectedChartDesignRange}
                  >
                    {selectedChartDesignRange || "—"}
                  </output>
                </label>
                <div className={spreadsheetHomeStyles.chartFormatDataActions}>
                  <button
                    type="button"
                    {...visibleCommandDomStateProps(
                      "XLSCTL.CONTEXT.CHART.SELECT_DATA",
                    )}
                    className={
                      spreadsheetHomeStyles.chartFormatPrimaryAction
                    }
                    onClick={() =>
                      void executeVisibleCommand(
                        "XLSCTL.CONTEXT.CHART.SELECT_DATA",
                      )
                    }
                    data-testid="spreadsheet-chart-format-select-data"
                  >
                    <OfficeSemanticIcon
                      semanticId="chart.selectData"
                      size={18}
                    />
                    {t("presentation.chartSelectData")}
                  </button>
                  <button
                    type="button"
                    {...chartDesignControlState(
                      "XLSCTL.CONTEXT.CHART.SWITCH_ROW_COLUMN",
                    ).dom}
                    className={
                      spreadsheetHomeStyles.chartFormatSecondaryAction
                    }
                    data-control-occurrence-id="format-switch-row-column"
                    data-control-surface="chart-format-pane"
                    data-testid="spreadsheet-chart-format-switch-row-column"
                    onClick={() =>
                      void executeChartDesignMutation(
                        "XLSCTL.CONTEXT.CHART.SWITCH_ROW_COLUMN",
                        String(sheetChartPlotsByRows(selectedChartDesignTarget)),
                      )
                    }
                  >
                    <OfficeSemanticIcon
                      semanticId="chart.switchRowColumn"
                      size={18}
                    />
                    {t("presentation.chartSwitchRowColumn")}
                  </button>
                </div>
              </div>
            </section>
          ) : null}
      </>
    );
    const parts = Children.toArray(options.props.children);
    return parts.length ? (
      <div className={spreadsheetHomeStyles.chartFormatOptions}>{parts}</div>
    ) : null;
  };
  return {
    renderSheetChartPaneOptions,
  };
}
