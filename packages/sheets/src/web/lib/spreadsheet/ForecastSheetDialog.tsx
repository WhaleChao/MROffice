"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Forecast Sheet dialog (数据 ▸ 预测工作表).
 *
 * Collects what Excel's wizard asks for — the timeline, the values, how far to
 * forecast, automatic/manual seasonality, the confidence level and where to
 * put the result — and hands it to the host. The maths lives in
 * `forecast-sheet.ts`.
 *
 * Display copy comes from the caller's active message catalogue.
 */

import { useEffect, useRef, useState } from "react";

import {
  officeDialogParts as dialogParts,
  spreadsheetDialogClassName,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { SpreadsheetCollapseDialogButton } from "../../components/notebook/contents/spreadsheet/SpreadsheetRangePicker";
import type { SpreadsheetRangePick } from "../../components/notebook/contents/spreadsheet/useSpreadsheetRangePick";

import {
  MAX_FORECAST_SEASONALITY,
  type ForecastDuplicateAggregation,
  type ForecastMissingPointPolicy,
} from "./forecast-sheet";

export type ForecastSheetSeasonalityRequest =
  { mode: "auto" } | { mode: "manual"; period: string };

export interface ForecastSheetRequest {
  timeline: string;
  values: string;
  periods: string;
  seasonality: ForecastSheetSeasonalityRequest;
  confidence: string;
  missingPoints: ForecastMissingPointPolicy;
  duplicateAggregation: ForecastDuplicateAggregation;
  chartType?: "line" | "column";
  destination: string;
}

export interface ForecastSheetDialogProps {
  open: boolean;
  translate: (key: string, values?: Record<string, string | number>) => string;
  onApply: (request: ForecastSheetRequest) => { ok: boolean; message: string };
  onClose: () => void;
  /** The selection's first and second columns when the dialog opens; Excel
   *  fills Timeline Range and Values Range from the selected data. */
  defaultTimeline?: string;
  defaultValues?: string;
  /** The window's RefEdit: both ranges get a Collapse Dialog button. */
  rangePick?: SpreadsheetRangePick;
}

export function ForecastSheetDialog({
  open,
  translate,
  onApply,
  onClose,
  defaultTimeline,
  defaultValues,
  rangePick,
}: ForecastSheetDialogProps) {
  const text = {
    title: translate("spreadsheet.helperCopy.forecastSheet.title"),
    timeline: translate("spreadsheet.helperCopy.forecastSheet.timeline"),
    values: translate("spreadsheet.helperCopy.forecastSheet.values"),
    periods: translate("spreadsheet.helperCopy.forecastSheet.periods"),
    seasonality: translate("spreadsheet.helperCopy.forecastSheet.seasonality"),
    seasonalityAuto: translate("spreadsheet.helperCopy.forecastSheet.seasonalityAuto"),
    seasonalityManual: translate("spreadsheet.helperCopy.forecastSheet.seasonalityManual"),
    seasonalityPeriod: translate("spreadsheet.helperCopy.forecastSheet.seasonalityPeriod"),
    seasonalityHint: translate("spreadsheet.helperCopy.forecastSheet.seasonalityHint"),
    confidence: translate("spreadsheet.helperCopy.forecastSheet.confidence"),
    missingPoints: translate("spreadsheet.helperCopy.forecastSheet.missingPoints"),
    missingInterpolate: translate("spreadsheet.helperCopy.forecastSheet.missingInterpolate"),
    missingZero: translate("spreadsheet.helperCopy.forecastSheet.missingZero"),
    missingHint: translate("spreadsheet.helperCopy.forecastSheet.missingHint"),
    duplicateAggregation: translate("spreadsheet.helperCopy.forecastSheet.duplicateAggregation"),
    chartType: translate("spreadsheet.helperCopy.forecastSheet.chartType"),
    chartLine: translate("spreadsheet.helperCopy.forecastSheet.chartLine"),
    chartColumn: translate("spreadsheet.helperCopy.forecastSheet.chartColumn"),
    average: translate("spreadsheet.helperCopy.forecastSheet.average"),
    sum: translate("spreadsheet.helperCopy.forecastSheet.sum"),
    count: translate("spreadsheet.helperCopy.forecastSheet.count"),
    counta: translate("spreadsheet.helperCopy.forecastSheet.counta"),
    min: translate("spreadsheet.helperCopy.forecastSheet.min"),
    max: translate("spreadsheet.helperCopy.forecastSheet.max"),
    median: translate("spreadsheet.helperCopy.forecastSheet.median"),
    destination: translate("spreadsheet.helperCopy.forecastSheet.destination"),
    apply: translate("spreadsheet.helperCopy.forecastSheet.apply"),
    close: translate("spreadsheet.helperCopy.forecastSheet.close"),
    hint: translate("spreadsheet.helperCopy.forecastSheet.hint"),
  };
  const [timeline, setTimeline] = useState("");
  const [values, setValues] = useState("");
  const [periods, setPeriods] = useState("6");
  const [seasonalityMode, setSeasonalityMode] = useState<"auto" | "manual">(
    "auto",
  );
  const [seasonalityPeriod, setSeasonalityPeriod] = useState("12");
  const [confidence, setConfidence] = useState("0.95");
  const [missingPoints, setMissingPoints] =
    useState<ForecastMissingPointPolicy>("interpolate");
  const [duplicateAggregation, setDuplicateAggregation] =
    useState<ForecastDuplicateAggregation>("average");
  const [chartType, setChartType] = useState<"line" | "column">("line");
  const [destination, setDestination] = useState(
    () => translate("spreadsheet.helperCopy.forecastSheet.title"),
  );
  const [status, setStatus] = useState<{ ok: boolean; message: string } | null>(
    null,
  );

  // The boxes start from the selection the dialog opened on. The defaults
  // follow the live selection, so a range picked on the sheet (or any
  // selection change while open) must not refill them.
  const seededRef = useRef(false);
  useEffect(() => {
    if (!open) {
      seededRef.current = false;
      return;
    }
    if (seededRef.current) return;
    seededRef.current = true;
    setStatus(null);
    setTimeline(defaultTimeline ?? "");
    setValues(defaultValues ?? "");
  }, [open, defaultTimeline, defaultValues]);

  if (!open) return null;
  // Collapsed to the RefEdit bar while one of its ranges is picked.
  if (rangePick?.picking("spreadsheet-forecast-")) return null;

  const pickButton = (
    id: "timeline" | "values",
    label: string,
    value: string,
    write: (text: string) => void,
  ) =>
    rangePick ? (
      <SpreadsheetCollapseDialogButton
        label={translate("spreadsheet.cfDialog.collapse")}
        onClick={() =>
          rangePick.start({
            id: `spreadsheet-forecast-${id}`,
            dialogLabel: text.title,
            inputLabel: label,
            style: "relative",
            value,
            onChange: (next) => {
              write(next);
              setStatus(null);
            },
          })
        }
        testId={`spreadsheet-forecast-${id}-pick`}
      />
    ) : null;

  return (
    <div
      className={spreadsheetDialogClassName("md")}
      role="dialog"
      aria-modal="true"
      aria-label={text.title}
      data-testid="spreadsheet-forecast-sheet-dialog"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          onClose();
        }
      }}
    >
      <span className={dialogParts.title}>{text.title}</span>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.timeline}</span>
        <span style={{ display: "flex", gap: 4 }}>
          <input
            value={timeline}
            onChange={(event) => setTimeline(event.target.value)}
            placeholder="A2:A20"
            data-testid="spreadsheet-forecast-timeline"
            style={{ flex: 1, minWidth: 0 }}
          />
          {pickButton("timeline", text.timeline, timeline, setTimeline)}
        </span>
      </label>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.chartType}</span>
        <select
          value={chartType}
          onChange={(event) => {
            setChartType(event.target.value as "line" | "column");
            setStatus(null);
          }}
          data-testid="spreadsheet-forecast-chart-type"
        >
          <option value="line">{text.chartLine}</option>
          <option value="column">{text.chartColumn}</option>
        </select>
      </label>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.values}</span>
        <span style={{ display: "flex", gap: 4 }}>
          <input
            value={values}
            onChange={(event) => setValues(event.target.value)}
            placeholder="B2:B20"
            data-testid="spreadsheet-forecast-values"
            style={{ flex: 1, minWidth: 0 }}
          />
          {pickButton("values", text.values, values, setValues)}
        </span>
      </label>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.periods}</span>
        <input
          value={periods}
          onChange={(event) => setPeriods(event.target.value)}
          inputMode="numeric"
          data-testid="spreadsheet-forecast-periods"
        />
      </label>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.seasonality}</span>
        <span style={{ display: "flex", gap: 6 }}>
          <select
            value={seasonalityMode}
            onChange={(event) => {
              setSeasonalityMode(event.target.value as "auto" | "manual");
              setStatus(null);
            }}
            aria-label={text.seasonality}
            data-testid="spreadsheet-forecast-seasonality-mode"
          >
            <option value="auto">{text.seasonalityAuto}</option>
            <option value="manual">{text.seasonalityManual}</option>
          </select>
          {seasonalityMode === "manual" ? (
            <input
              type="number"
              value={seasonalityPeriod}
              onChange={(event) => {
                setSeasonalityPeriod(event.target.value);
                setStatus(null);
              }}
              inputMode="numeric"
              min={2}
              max={MAX_FORECAST_SEASONALITY}
              step={1}
              aria-label={text.seasonalityPeriod}
              data-testid="spreadsheet-forecast-seasonality-period"
              style={{ minWidth: 0, width: 88 }}
            />
          ) : null}
        </span>
        <span className={dialogParts.hint}>{text.seasonalityHint}</span>
      </label>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.confidence}</span>
        <input
          value={confidence}
          onChange={(event) => setConfidence(event.target.value)}
          inputMode="decimal"
          data-testid="spreadsheet-forecast-confidence"
        />
      </label>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.missingPoints}</span>
        <select
          value={missingPoints}
          onChange={(event) => {
            setMissingPoints(event.target.value as ForecastMissingPointPolicy);
            setStatus(null);
          }}
          data-testid="spreadsheet-forecast-missing-points"
        >
          <option value="interpolate">{text.missingInterpolate}</option>
          <option value="zero">{text.missingZero}</option>
        </select>
        <span className={dialogParts.hint}>{text.missingHint}</span>
      </label>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.duplicateAggregation}</span>
        <select
          value={duplicateAggregation}
          onChange={(event) => {
            setDuplicateAggregation(
              event.target.value as ForecastDuplicateAggregation,
            );
            setStatus(null);
          }}
          data-testid="spreadsheet-forecast-duplicate-aggregation"
        >
          <option value="average">{text.average}</option>
          <option value="sum">{text.sum}</option>
          <option value="count">{text.count}</option>
          <option value="counta">{text.counta}</option>
          <option value="min">{text.min}</option>
          <option value="max">{text.max}</option>
          <option value="median">{text.median}</option>
        </select>
      </label>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.destination}</span>
        <input
          value={destination}
          onChange={(event) => setDestination(event.target.value)}
          maxLength={31}
          data-testid="spreadsheet-forecast-destination"
        />
      </label>
      <span className={dialogParts.hint}>{text.hint}</span>
      {status ? (
        <span
          data-testid="spreadsheet-forecast-status"
          className={dialogParts.message}
          data-tone={status.ok ? undefined : "error"}
        >
          {status.message}
        </span>
      ) : null}
      <div className={dialogParts.footer}>
        <button
          type="button"
          data-testid="spreadsheet-forecast-close"
          onClick={onClose}
          className={dialogParts.button}
        >
          {text.close}
        </button>
        <button
          type="button"
          data-testid="spreadsheet-forecast-apply"
          onClick={() => {
            const result = onApply({
              timeline,
              values,
              periods,
              seasonality:
                seasonalityMode === "auto"
                  ? { mode: "auto" }
                  : { mode: "manual", period: seasonalityPeriod },
              confidence,
              missingPoints,
              duplicateAggregation,
              chartType,
              destination,
            });
            if (result.ok) onClose();
            else setStatus(result);
          }}
          className={dialogParts.primaryButton}
        >
          {text.apply}
        </button>
      </div>
    </div>
  );
}

export default ForecastSheetDialog;
