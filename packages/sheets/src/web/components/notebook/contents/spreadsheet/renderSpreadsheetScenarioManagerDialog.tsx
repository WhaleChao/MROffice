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
import { formatScenarioChangingCells, type Scenario } from "../../../../lib/spreadsheet/scenarios";
import { type ScenarioFormDraft } from "./spreadsheetWindowLocalTypes";
import { type Dispatch, type SetStateAction } from "react";
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

/** What renderSpreadsheetScenarioManagerDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetScenarioManagerDialogInputs {
  deleteScenario: (name: string) => void;
  openScenarioForm: (mode: "add" | "edit") => void;
  openScenarioMerge: () => void;
  openScenarioSummary: () => void;
  scenarioForm: ScenarioFormDraft | null;
  scenarioManagerFocusRef: RefObject<HTMLButtonElement | null>;
  scenarioMergeDraft: { sheetId: string; sourceSheetId: string | null; } | null;
  scenarioOpen: boolean;
  scenarioStatus: { kind: "success" | "error"; text: string; } | null;
  scenarioSummaryDraft: { sheetId: string; resultCells: string; reportType: "summary" | "pivot"; error: string | null; } | null;
  scenarioValuesDraft: { form: ScenarioFormDraft; refs: readonly string[]; values: readonly string[]; } | null;
  selectedScenario: Scenario;
  selectedScenarioLocked: boolean;
  setScenarioOpen: Dispatch<SetStateAction<boolean>>;
  setScenarioSelectedName: Dispatch<SetStateAction<string | null>>;
  setScenarioStatus: Dispatch<SetStateAction<{ kind: "success" | "error"; text: string; } | null>>;
  showScenario: (name: string) => void;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  visibleScenarios: Scenario[];
  whatIfAnalysisReturnFocusTarget: () => HTMLElement | null;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetScenarioManagerDialog({
  deleteScenario,
  openScenarioForm,
  openScenarioMerge,
  openScenarioSummary,
  scenarioForm,
  scenarioManagerFocusRef,
  scenarioMergeDraft,
  scenarioOpen,
  scenarioStatus,
  scenarioSummaryDraft,
  scenarioValuesDraft,
  selectedScenario,
  selectedScenarioLocked,
  setScenarioOpen,
  setScenarioSelectedName,
  setScenarioStatus,
  showScenario,
  t,
  visibleScenarios,
  whatIfAnalysisReturnFocusTarget,
}: RenderSpreadsheetScenarioManagerDialogInputs) {
  return (
      <OfficeDialogShell
        open={
          scenarioOpen &&
          scenarioForm === null &&
          scenarioValuesDraft === null &&
          scenarioMergeDraft === null &&
          scenarioSummaryDraft === null
        }
        title={t("spreadsheet.scenarioMenu")}
        description={null}
        testIdPrefix="spreadsheet-scenario"
        defaultActionTestId="spreadsheet-scenario-close"
        defaultActionLabel={t("common.close")}
        cancelActionLabel={t("common.cancel")}
        showCancelAction={false}
        onDefaultAction={() => {
          setScenarioOpen(false);
          setScenarioStatus(null);
        }}
        onCancel={() => {
          setScenarioOpen(false);
          setScenarioStatus(null);
        }}
        initialFocusRef={scenarioManagerFocusRef}
        returnFocusTarget={() =>
          // A sub-dialog is taking over: its own focus trap holds.
          scenarioForm !== null ||
          scenarioValuesDraft !== null ||
          scenarioMergeDraft !== null ||
          scenarioSummaryDraft !== null
            ? null
            : whatIfAnalysisReturnFocusTarget()
        }
        footerStart={
          <button
            type="button"
            onClick={() => {
              if (selectedScenario) showScenario(selectedScenario.name);
            }}
            disabled={selectedScenario === null}
            data-testid="spreadsheet-scenario-show"
            className={dialogParts.button}
          >
            <SpreadsheetDialogIcon name="show" />
            {t("spreadsheet.scenarioShow")}
          </button>
        }
        className="sm:max-w-[460px]"
      >
        <span className={dialogParts.label}>
          {t("spreadsheet.scenarioListLabel")}
        </span>
        <span
          style={{
            display: "grid",
            gridTemplateColumns: "minmax(0, 1fr) auto",
            gap: 8,
            alignItems: "start",
          }}
        >
          {visibleScenarios.length > 0 ? (
            <span
              role="listbox"
              aria-label={t("spreadsheet.scenarioListLabel")}
              data-testid="spreadsheet-scenario-list"
              className={dialogParts.list}
              style={{ maxHeight: 160, minHeight: 96, overflowY: "auto" }}
            >
              {visibleScenarios.map((scenario, index) => {
                const selected = scenario === selectedScenario;
                return (
                  <button
                    key={scenario.name}
                    ref={selected ? scenarioManagerFocusRef : undefined}
                    type="button"
                    role="option"
                    aria-selected={selected}
                    tabIndex={selected ? 0 : -1}
                    onClick={() => setScenarioSelectedName(scenario.name)}
                    onDoubleClick={() => showScenario(scenario.name)}
                    onKeyDown={(event) => {
                      const step =
                        event.key === "ArrowDown"
                          ? 1
                          : event.key === "ArrowUp"
                            ? -1
                            : 0;
                      if (step === 0) return;
                      event.preventDefault();
                      const next = visibleScenarios[index + step];
                      if (!next) return;
                      setScenarioSelectedName(next.name);
                      const sibling =
                        step > 0
                          ? event.currentTarget.nextElementSibling
                          : event.currentTarget.previousElementSibling;
                      if (sibling instanceof HTMLElement) sibling.focus();
                    }}
                    data-testid="spreadsheet-scenario-item"
                    data-name={scenario.name}
                    className={dialogParts.listItem}
                    title={scenario.name}
                  >
                    {scenario.name}
                  </button>
                );
              })}
            </span>
          ) : (
            <span
              data-testid="spreadsheet-scenario-empty"
              className={dialogParts.hint}
              style={{ minHeight: 96 }}
            >
              {t("spreadsheet.scenarioNone")}
            </span>
          )}
          <span style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            <button
              ref={
                visibleScenarios.length === 0
                  ? scenarioManagerFocusRef
                  : undefined
              }
              type="button"
              onClick={() => openScenarioForm("add")}
              data-testid="spreadsheet-scenario-add"
              className={dialogParts.button}
            >
              <SpreadsheetDialogIcon name="add" />
              {t("spreadsheet.scenarioAddEllipsis")}
            </button>
            <button
              type="button"
              onClick={() => {
                if (selectedScenario) deleteScenario(selectedScenario.name);
                setScenarioSelectedName(null);
              }}
              disabled={selectedScenario === null || selectedScenarioLocked}
              data-testid="spreadsheet-scenario-delete"
              className={dialogParts.button}
            >
              <SpreadsheetDialogIcon name="delete" />
              {t("spreadsheet.scenarioDeleteButton")}
            </button>
            <button
              type="button"
              onClick={() => openScenarioForm("edit")}
              disabled={selectedScenario === null || selectedScenarioLocked}
              data-testid="spreadsheet-scenario-edit"
              className={dialogParts.button}
            >
              <SpreadsheetDialogIcon name="edit" />
              {t("spreadsheet.scenarioEditEllipsis")}
            </button>
            <button
              type="button"
              onClick={openScenarioMerge}
              data-testid="spreadsheet-scenario-merge"
              className={dialogParts.button}
            >
              <SpreadsheetDialogIcon name="merge" />
              {t("spreadsheet.scenarioMergeEllipsis")}
            </button>
            <button
              type="button"
              onClick={openScenarioSummary}
              disabled={visibleScenarios.length === 0}
              data-testid="spreadsheet-scenario-summary"
              className={dialogParts.button}
            >
              <SpreadsheetDialogIcon name="summary" />
              {t("spreadsheet.scenarioSummaryEllipsis")}
            </button>
          </span>
        </span>
        {selectedScenario ? (
          <span
            style={{
              display: "grid",
              gridTemplateColumns: "auto minmax(0, 1fr)",
              gap: "4px 8px",
            }}
          >
            <span className={dialogParts.label}>
              {t("spreadsheet.scenarioChangingCells")}
            </span>
            <span
              data-testid="spreadsheet-scenario-changing-cells"
              style={{ wordBreak: "break-all" }}
            >
              {formatScenarioChangingCells(selectedScenario.cells)}
            </span>
            <span className={dialogParts.label}>
              {t("spreadsheet.scenarioComment")}
            </span>
            <span
              data-testid="spreadsheet-scenario-comment"
              style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}
            >
              {selectedScenario.comment ?? ""}
            </span>
          </span>
        ) : null}
        {scenarioStatus ? (
          <span
            role="status"
            data-testid="spreadsheet-scenario-status"
            data-status={scenarioStatus.kind}
            className={dialogParts.message}
            data-tone={scenarioStatus.kind === "success" ? "success" : "error"}
          >
            {scenarioStatus.text}
          </span>
        ) : null}
      </OfficeDialogShell>
  );
}
