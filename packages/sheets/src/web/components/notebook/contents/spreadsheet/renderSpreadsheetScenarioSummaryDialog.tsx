"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { type Dispatch, type SetStateAction } from "react";
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import { type SpreadsheetRangePick } from "./useSpreadsheetRangePick";

/** What renderSpreadsheetScenarioSummaryDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetScenarioSummaryDialogInputs {
  rangePick: SpreadsheetRangePick;
  commitScenarioSummary: () => Promise<void>;
  scenarioOpen: boolean;
  scenarioPivotUnavailableReason: string | null;
  scenarioSummaryDraft: { sheetId: string; resultCells: string; reportType: "summary" | "pivot"; error: string | null; } | null;
  scenarioSummaryFocusRef: RefObject<HTMLInputElement | null>;
  setScenarioSummaryDraft: Dispatch<SetStateAction<{ sheetId: string; resultCells: string; reportType: "summary" | "pivot"; error: string | null; } | null>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetScenarioSummaryDialog({
  rangePick,
  commitScenarioSummary,
  scenarioOpen,
  scenarioPivotUnavailableReason,
  scenarioSummaryDraft,
  scenarioSummaryFocusRef,
  setScenarioSummaryDraft,
  t,
}: RenderSpreadsheetScenarioSummaryDialogInputs) {
  // Collapsed while a range of it is picked on the sheet.
  if (rangePick.picking("spreadsheet-scenario-summary")) return null;
  return (
      <OfficeDialogShell
        open={scenarioOpen && scenarioSummaryDraft !== null}
        title={t("spreadsheet.scenarioSummaryTitle")}
        description={null}
        testIdPrefix="spreadsheet-scenario-summary"
        defaultActionLabel={t("spreadsheet.fmtOk")}
        cancelActionLabel={t("common.cancel")}
        onDefaultAction={commitScenarioSummary}
        onCancel={() => setScenarioSummaryDraft(null)}
        initialFocusRef={scenarioSummaryFocusRef}
        className="sm:max-w-[360px]"
      >
        <span className={dialogParts.sectionTitle}>
          {t("spreadsheet.scenarioSummaryReportType")}
        </span>
        <span className={dialogParts.choices} role="radiogroup">
          <label className={dialogParts.choice}>
            <input
              type="radio"
              name="spreadsheet-scenario-summary-type"
              checked={scenarioSummaryDraft?.reportType !== "pivot"}
              onChange={() =>
                setScenarioSummaryDraft((current) =>
                  current ? { ...current, reportType: "summary", error: null } : current,
                )
              }
              data-testid="spreadsheet-scenario-summary-type-summary"
            />
            <span>{t("spreadsheet.scenarioSummaryStandard")}</span>
          </label>
          <label
            className={dialogParts.choice}
            title={scenarioPivotUnavailableReason ?? undefined}
            data-disabled-reason={scenarioPivotUnavailableReason ?? undefined}
          >
            <input
              type="radio"
              name="spreadsheet-scenario-summary-type"
              checked={scenarioSummaryDraft?.reportType === "pivot"}
              // A choice already made stays chosen while the box is edited;
              // OK then says what the report needs.
              disabled={
                scenarioPivotUnavailableReason !== null &&
                scenarioSummaryDraft?.reportType !== "pivot"
              }
              onChange={() =>
                setScenarioSummaryDraft((current) =>
                  current ? { ...current, reportType: "pivot", error: null } : current,
                )
              }
              aria-describedby={
                scenarioPivotUnavailableReason
                  ? "spreadsheet-scenario-summary-pivot-reason"
                  : undefined
              }
              data-testid="spreadsheet-scenario-summary-type-pivot"
            />
            <span>{t("spreadsheet.scenarioSummaryPivot")}</span>
          </label>
          {scenarioPivotUnavailableReason ? (
            <span
              id="spreadsheet-scenario-summary-pivot-reason"
              className={dialogParts.hint}
            >
              {scenarioPivotUnavailableReason}
            </span>
          ) : null}
        </span>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.scenarioSummaryResultCells")}
          </span>
          <span style={{ display: "flex", gap: 4 }}>
            <input
              ref={scenarioSummaryFocusRef}
              value={scenarioSummaryDraft?.resultCells ?? ""}
              onChange={(event) => {
                const resultCells = event.target.value;
                setScenarioSummaryDraft((current) =>
                  current ? { ...current, resultCells, error: null } : current,
                );
              }}
              aria-invalid={scenarioSummaryDraft?.error ? true : undefined}
              aria-describedby={
                scenarioSummaryDraft?.error
                  ? "spreadsheet-scenario-summary-error"
                  : undefined
              }
              data-testid="spreadsheet-scenario-summary-result-cells"
              style={{ flex: 1, minWidth: 0 }}
            />
            <SpreadsheetCollapseDialogButton
              label={t("spreadsheet.cfDialog.collapse")}
              onClick={() =>
                rangePick.start({
                  id: "spreadsheet-scenario-summary-result-cells",
                  dialogLabel: t("spreadsheet.scenarioSummaryTitle"),
                  inputLabel: t("spreadsheet.scenarioSummaryResultCells"),
                  style: "absolute",
                  value: String(scenarioSummaryDraft?.resultCells ?? ""),
                  onChange: (text: string) => {
                    const resultCells = text;
                    setScenarioSummaryDraft((current) =>
                      current ? { ...current, resultCells, error: null } : current,
                    );
                  },
                })
              }
              testId="spreadsheet-scenario-summary-result-cells-pick"
            />
          </span>
        </label>
        {scenarioSummaryDraft?.error ? (
          <span
            id="spreadsheet-scenario-summary-error"
            role="alert"
            className={dialogParts.error}
            data-testid="spreadsheet-scenario-summary-error"
          >
            {scenarioSummaryDraft.error}
          </span>
        ) : null}
      </OfficeDialogShell>
  );
}
