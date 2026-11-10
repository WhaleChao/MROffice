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
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import { MAX_SCENARIO_NAME_LENGTH, MAX_SCENARIO_COMMENT_LENGTH } from "../../../../lib/spreadsheet/scenarios";
import { type ScenarioFormDraft } from "./spreadsheetWindowLocalTypes";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetScenarioEditDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetScenarioEditDialogInputs {
  scenarioCellsPicking: { r1: number; c1: number; r2: number; c2: number; } | null;
  scenarioForm: ScenarioFormDraft | null;
  scenarioNameInputRef: RefObject<HTMLInputElement | null>;
  scenarioOpen: boolean;
  scenarioValuesDraft: { form: ScenarioFormDraft; refs: readonly string[]; values: readonly string[]; } | null;
  setScenarioCellsPicking: Dispatch<SetStateAction<{ r1: number; c1: number; r2: number; c2: number; } | null>>;
  setScenarioForm: Dispatch<SetStateAction<ScenarioFormDraft | null>>;
  startScenarioCellsPick: () => void;
  submitScenarioForm: () => void;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetScenarioEditDialog({
  scenarioCellsPicking,
  scenarioForm,
  scenarioNameInputRef,
  scenarioOpen,
  scenarioValuesDraft,
  setScenarioCellsPicking,
  setScenarioForm,
  startScenarioCellsPick,
  submitScenarioForm,
  t,
}: RenderSpreadsheetScenarioEditDialogInputs) {
  return (
      <OfficeDialogShell
        open={
          scenarioOpen &&
          scenarioForm !== null &&
          scenarioValuesDraft === null &&
          scenarioCellsPicking === null
        }
        title={
          scenarioForm?.mode === "edit"
            ? t("spreadsheet.scenarioEditTitle")
            : t("spreadsheet.scenarioAddTitle")
        }
        description={null}
        testIdPrefix="spreadsheet-scenario-form"
        defaultActionLabel={t("spreadsheet.fmtOk")}
        cancelActionLabel={t("common.cancel")}
        onDefaultAction={submitScenarioForm}
        onCancel={() => {
          setScenarioForm(null);
          setScenarioCellsPicking(null);
        }}
        initialFocusRef={scenarioNameInputRef}
        className="sm:max-w-[420px]"
      >
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.scenarioNameField")}
          </span>
          <input
            ref={scenarioNameInputRef}
            value={scenarioForm?.name ?? ""}
            maxLength={MAX_SCENARIO_NAME_LENGTH}
            onChange={(event) => {
              const name = event.target.value;
              setScenarioForm((current) =>
                current ? { ...current, name, error: null } : current,
              );
            }}
            data-testid="spreadsheet-scenario-name"
          />
        </label>
        <span className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.scenarioChangingCells")}
          </span>
          <span style={{ display: "flex", gap: 4 }}>
            <input
              value={scenarioForm?.changingCells ?? ""}
              aria-label={t("spreadsheet.scenarioChangingCells")}
              onChange={(event) => {
                const changingCells = event.target.value;
                setScenarioForm((current) =>
                  current
                    ? { ...current, changingCells, error: null }
                    : current,
                );
              }}
              data-testid="spreadsheet-scenario-cells"
              style={{ flex: 1, minWidth: 0 }}
            />
            <SpreadsheetCollapseDialogButton
              label={t("spreadsheet.cfDialog.collapse")}
              onClick={startScenarioCellsPick}
              testId="spreadsheet-scenario-cells-collapse"
            />
          </span>
          <span className={dialogParts.hint}>
            {t("spreadsheet.scenarioChangingCellsHint")}
          </span>
        </span>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.scenarioComment")}
          </span>
          <textarea
            value={scenarioForm?.comment ?? ""}
            maxLength={MAX_SCENARIO_COMMENT_LENGTH}
            rows={3}
            onChange={(event) => {
              const comment = event.target.value;
              setScenarioForm((current) =>
                current ? { ...current, comment } : current,
              );
            }}
            data-testid="spreadsheet-scenario-comment-input"
          />
        </label>
        <fieldset
          style={{ border: 0, margin: 0, padding: 0, display: "grid", gap: 4 }}
        >
          <legend className={dialogParts.label}>
            {t("spreadsheet.scenarioProtection")}
          </legend>
          <label className={dialogParts.check}>
            <input
              type="checkbox"
              checked={scenarioForm?.locked ?? true}
              onChange={(event) => {
                const locked = event.target.checked;
                setScenarioForm((current) =>
                  current ? { ...current, locked } : current,
                );
              }}
              data-testid="spreadsheet-scenario-prevent-changes"
            />
            {t("spreadsheet.scenarioPreventChanges")}
          </label>
          <label className={dialogParts.check}>
            <input
              type="checkbox"
              checked={scenarioForm?.hidden ?? false}
              onChange={(event) => {
                const hidden = event.target.checked;
                setScenarioForm((current) =>
                  current ? { ...current, hidden } : current,
                );
              }}
              data-testid="spreadsheet-scenario-hide"
            />
            {t("spreadsheet.scenarioHide")}
          </label>
        </fieldset>
        {scenarioForm?.error ? (
          <span
            role="alert"
            data-testid="spreadsheet-scenario-form-error"
            className={dialogParts.message}
            data-tone="error"
          >
            {scenarioForm.error}
          </span>
        ) : null}
      </OfficeDialogShell>
  );
}
