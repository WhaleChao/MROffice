"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { type ForecastSheetRequest } from "../../../../lib/spreadsheet/ForecastSheetDialog";
import { forecastCellMapText, forecastSheetNameValid, forecastTableRows, type ForecastOutputRequest } from "../../../../lib/spreadsheet/forecast-sheet-output";
import { parseA1Range } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import { type SheetModel } from "./spreadsheetSheetModel";

/** What useSpreadsheetRunForecastSheet reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetRunForecastSheetInputs {
  forecastOutputWriterRef: RefObject<((request: ForecastOutputRequest) => "ok" | "sheet-limit" | "invalid-name" | "duplicate-name" | "workbook-protected") | null>;
  locale: string;
  sheet: SheetModel;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  workbookStructureProtectedReason: string;
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function useSpreadsheetRunForecastSheet({
  forecastOutputWriterRef,
  locale,
  sheet,
  t,
  workbookStructureProtectedReason,
}: SpreadsheetRunForecastSheetInputs) {
  const runForecastSheet = useCallback(
    (request: ForecastSheetRequest): { ok: boolean; message: string } => {
      const zh = locale.toLowerCase().startsWith("zh");
      const invalid = (what: string) => ({
        ok: false,
        message: t("spreadsheet.dynamicCopy.invalidForecastField", { field: what }),
      });
      const destination = request.destination.trim();
      if (!parseA1Range(request.timeline.trim())) return invalid(t("spreadsheet.uiCopy.timelineRange"));
      if (!parseA1Range(request.values.trim())) return invalid(t("spreadsheet.uiCopy.valuesRange"));
      if (!forecastSheetNameValid(destination)) {
        return invalid(t("spreadsheet.uiCopy.outputSheetName"));
      }
      const manualSeasonality =
        request.seasonality.mode === "manual"
          ? Number(request.seasonality.period.trim())
          : null;
      // A manual 1 must not be reinterpreted as automatic detection; pass an
      // invalid sentinel so the model returns its precise fail-closed error.
      const seasonality =
        manualSeasonality === null
          ? 1
          : manualSeasonality >= 2
            ? manualSeasonality
            : Number.NaN;
      // The shared reader + builder (lib/spreadsheet/forecast-sheet-output),
      // which the AI's spreadsheet.data.forecast_sheet command runs too.
      const built = forecastTableRows(forecastCellMapText(sheet.cells), {
        timeline: request.timeline,
        values: request.values,
        periods: Number(request.periods.trim()),
        seasonality,
        missingPoints: request.missingPoints,
        duplicateAggregation: request.duplicateAggregation,
        // Accept both "0.95" and "95": the dialog stays forgiving while the
        // model stays strict about what a confidence level may be.
        confidence: (() => {
          const typed = Number(request.confidence.trim());
          if (!Number.isFinite(typed)) return 0.95;
          return typed > 1 ? typed / 100 : typed;
        })(),
        // These headers become persisted worksheet data, not UI catalogue labels.
        language: zh ? "zh" : "en",
        headers: [
          t("spreadsheet.forecastSheet.header.timeline"),
          t("spreadsheet.forecastSheet.header.values"),
          t("spreadsheet.forecastSheet.header.forecast"),
          t("spreadsheet.forecastSheet.header.lower"),
          t("spreadsheet.forecastSheet.header.upper"),
        ],
      });
      if (!built.ok) {
        if (built.error === "timeline-range") return invalid(t("spreadsheet.uiCopy.timelineRange"));
        if (built.error === "values-range") return invalid(t("spreadsheet.uiCopy.valuesRange"));
        const reason: Record<string, string> = {
          "too-few-points": "spreadsheet.dynamicCopy.forecastError.too-few-points",
          "mismatched-lengths": "spreadsheet.dynamicCopy.forecastError.mismatched-lengths",
          "invalid-periods": "spreadsheet.dynamicCopy.forecastError.invalid-periods",
          "invalid-confidence": "spreadsheet.dynamicCopy.forecastError.invalid-confidence",
          "invalid-seasonality": "spreadsheet.dynamicCopy.forecastError.invalid-seasonality",
          "invalid-timeline": "spreadsheet.dynamicCopy.forecastError.invalid-timeline",
          "too-many-missing-points": "spreadsheet.dynamicCopy.forecastError.too-many-missing-points",
          "unfillable-missing-points": "spreadsheet.dynamicCopy.forecastError.unfillable-missing-points",
          "not-forecastable": "spreadsheet.dynamicCopy.forecastError.not-forecastable"
        };
        return { ok: false, message: t(reason[built.error] ?? "spreadsheet.dynamicCopy.forecastError.failed") };
      }
      const output = built.rows;
      const writeResult =
        forecastOutputWriterRef.current?.({
          sheetName: destination,
          chartType: request.chartType ?? "line",
          rows: output,
          snapshot: built.snapshot,
          sourceSheetName: sheet.name,
          sourceRanges: [request.timeline.trim(), request.values.trim()],
        }) ?? "sheet-limit";
      if (writeResult !== "ok") {
        const reason =
          writeResult === "workbook-protected"
            ? workbookStructureProtectedReason
            : writeResult === "duplicate-name"
            ? t("spreadsheet.uiCopy.thatOutputSheetNameAlreadyExists")
            : writeResult === "invalid-name"
              ? t("spreadsheet.uiCopy.thatOutputSheetNameIsInvalid")
              : t("spreadsheet.uiCopy.theWorkbookHasReachedItsSheetLimit");
        return {
          ok: false,
          message: reason,
        };
      }
      return {
        ok: true,
        message: t("spreadsheet.dynamicCopy.forecastCreated", { count: built.forecastCount, sheet: destination }),
      };
    },
    [locale, sheet, workbookStructureProtectedReason, t],
  );
  return {
    runForecastSheet,
  };
}
