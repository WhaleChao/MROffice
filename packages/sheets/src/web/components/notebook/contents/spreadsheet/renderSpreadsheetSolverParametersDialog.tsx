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
import { type SolverDialogMethod, type SolverDialogOp, type SolverDialogConstraint } from "../../../../lib/spreadsheet/solver-model";
import { type SheetModel } from "./spreadsheetSheetModel";
import { type Dispatch, type SetStateAction } from "react";
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import { type SpreadsheetRangePick } from "./useSpreadsheetRangePick";
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

/** What renderSpreadsheetSolverParametersDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetSolverParametersDialogInputs {
  rangePick: SpreadsheetRangePick;
  applySolverResult: () => void;
  openSolverConstraintEditor: (index: number | null) => void;
  removeSolverConstraint: (index: number) => void;
  runSolver: () => Promise<void>;
  setSolverChangingRange: Dispatch<SetStateAction<string>>;
  setSolverGoalKind: Dispatch<SetStateAction<"min" | "max" | "value">>;
  setSolverGoalValue: Dispatch<SetStateAction<string>>;
  setSolverMethod: Dispatch<SetStateAction<SolverDialogMethod>>;
  setSolverNonNegative: Dispatch<SetStateAction<boolean>>;
  setSolverObjectiveCell: Dispatch<SetStateAction<string>>;
  setSolverOpen: Dispatch<SetStateAction<boolean>>;
  setSolverPending: Dispatch<SetStateAction<{ sourceSheet: SheetModel; values: readonly { row: number; col: number; value: string; }[]; } | null>>;
  setSolverSelectedConstraint: Dispatch<SetStateAction<number | null>>;
  setSolverStatus: Dispatch<SetStateAction<{ kind: "success" | "error"; text: string; } | null>>;
  setSolverSubDialog: Dispatch<SetStateAction<{ kind: "constraint"; index: number | null; cellRef: string; op: SolverDialogOp; bound: string; } | { kind: "options"; } | { kind: "model"; range: string; } | { kind: "reset"; } | null>>;
  solverChangingRange: string;
  solverConstraints: SolverDialogConstraint[];
  solverConstraintText: (constraint: SolverDialogConstraint) => string;
  solverGoalKind: "min" | "max" | "value";
  solverGoalValue: string;
  solverMethod: SolverDialogMethod;
  solverNonNegative: boolean;
  solverObjectiveCell: string;
  solverObjectiveInputRef: RefObject<HTMLInputElement | null>;
  solverOpen: boolean;
  solverPending: { sourceSheet: SheetModel; values: readonly { row: number; col: number; value: string; }[]; } | null;
  solverSelectedConstraint: number | null;
  solverStatus: { kind: "success" | "error"; text: string; } | null;
  solverSubDialog: { kind: "constraint"; index: number | null; cellRef: string; op: SolverDialogOp; bound: string; } | { kind: "options"; } | { kind: "model"; range: string; } | { kind: "reset"; } | null;
  solverTrialContinueRef: RefObject<boolean>;
  solverTrialOpen: boolean;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  whatIfAnalysisReturnFocusTarget: () => HTMLElement | null;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetSolverParametersDialog({
  rangePick,
  applySolverResult,
  openSolverConstraintEditor,
  removeSolverConstraint,
  runSolver,
  setSolverChangingRange,
  setSolverGoalKind,
  setSolverGoalValue,
  setSolverMethod,
  setSolverNonNegative,
  setSolverObjectiveCell,
  setSolverOpen,
  setSolverPending,
  setSolverSelectedConstraint,
  setSolverStatus,
  setSolverSubDialog,
  solverChangingRange,
  solverConstraints,
  solverConstraintText,
  solverGoalKind,
  solverGoalValue,
  solverMethod,
  solverNonNegative,
  solverObjectiveCell,
  solverObjectiveInputRef,
  solverOpen,
  solverPending,
  solverSelectedConstraint,
  solverStatus,
  solverSubDialog,
  solverTrialContinueRef,
  solverTrialOpen,
  t,
  whatIfAnalysisReturnFocusTarget,
}: RenderSpreadsheetSolverParametersDialogInputs) {
  // Collapsed while a range of it is picked on the sheet.
  if (rangePick.picking("spreadsheet-solver-")) return null;
  return (
      <OfficeDialogShell
        open={solverOpen && solverSubDialog === null}
        title={t("spreadsheet.uiCopy.solverParametersTitle")}
        description={null}
        testIdPrefix="spreadsheet-solver"
        defaultActionTestId="spreadsheet-solver-run"
        defaultActionData={{ "data-mode": solverPending ? "apply" : "solve" }}
        defaultActionLabel={t(
          solverPending
            ? "spreadsheet.goalSeekApplyResult"
            : "spreadsheet.solverRun",
        )}
        cancelActionLabel={t("spreadsheet.solverCancel")}
        onDefaultAction={solverPending ? applySolverResult : runSolver}
        onCancel={() => {
          setSolverOpen(false);
          setSolverPending(null);
          setSolverStatus(null);
        }}
        initialFocusRef={solverObjectiveInputRef}
        returnFocusTarget={whatIfAnalysisReturnFocusTarget}
        className="sm:max-w-[520px]"
      >
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.uiCopy.solverSetObjective")}
          </span>
          <span style={{ display: "flex", gap: 4 }}>
            <input
              ref={solverObjectiveInputRef}
              value={solverObjectiveCell}
              onChange={(event) => {
                setSolverObjectiveCell(event.target.value);
                setSolverPending(null);
                setSolverStatus(null);
              }}
              spellCheck={false}
              data-testid="spreadsheet-solver-objective"
              style={{ flex: 1, minWidth: 0 }}
            />
            <SpreadsheetCollapseDialogButton
              label={t("spreadsheet.cfDialog.collapse")}
              onClick={() =>
                rangePick.start({
                  id: "spreadsheet-solver-objective",
                  dialogLabel: t("spreadsheet.uiCopy.solverParametersTitle"),
                  inputLabel: t("spreadsheet.uiCopy.solverSetObjective"),
                  style: "absolute",
                  value: String(solverObjectiveCell),
                  onChange: (text: string) => {
                    setSolverObjectiveCell(text);
                    setSolverPending(null);
                    setSolverStatus(null);
                  },
                })
              }
              testId="spreadsheet-solver-objective-pick"
            />
          </span>
        </label>
        <fieldset
          className="m-0 flex flex-wrap items-center gap-x-4 gap-y-1 border-0 p-0"
          data-testid="spreadsheet-solver-goal-kind"
        >
          <legend className={dialogParts.label}>{t("spreadsheet.uiCopy.solverTo")}</legend>
          {(
            [
              ["max", "spreadsheet.solverGoalMax"],
              ["min", "spreadsheet.solverGoalMin"],
              ["value", "spreadsheet.uiCopy.solverValueOf"],
            ] as const
          ).map(([kind, label]) => (
            <label key={kind} className={dialogParts.check}>
              <input
                type="radio"
                name="spreadsheet-solver-goal"
                checked={solverGoalKind === kind}
                onChange={() => {
                  setSolverGoalKind(kind);
                  setSolverPending(null);
                  setSolverStatus(null);
                }}
                data-testid={`spreadsheet-solver-goal-kind-${kind}`}
              />
              <SpreadsheetDialogIcon name={kind === "max" ? "max" : kind === "min" ? "min" : "target"} />
              <span>{t(label)}</span>
            </label>
          ))}
          <input
            value={solverGoalValue}
            disabled={solverGoalKind !== "value"}
            onChange={(event) => {
              setSolverGoalValue(event.target.value);
              setSolverPending(null);
              setSolverStatus(null);
            }}
            aria-label={t("spreadsheet.uiCopy.solverValueOf")}
            inputMode="decimal"
            data-testid="spreadsheet-solver-goal-value"
            style={{ width: 96 }}
          />
        </fieldset>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.uiCopy.solverByChangingVariableCells")}
          </span>
          <span style={{ display: "flex", gap: 4 }}>
            <input
              value={solverChangingRange}
              onChange={(event) => {
                setSolverChangingRange(event.target.value);
                setSolverPending(null);
                setSolverStatus(null);
              }}
              placeholder={t("spreadsheet.solverRangePlaceholder")}
              spellCheck={false}
              data-testid="spreadsheet-solver-changing"
              style={{ flex: 1, minWidth: 0 }}
            />
            <SpreadsheetCollapseDialogButton
              label={t("spreadsheet.cfDialog.collapse")}
              onClick={() =>
                rangePick.start({
                  id: "spreadsheet-solver-changing",
                  dialogLabel: t("spreadsheet.uiCopy.solverParametersTitle"),
                  inputLabel: t("spreadsheet.uiCopy.solverByChangingVariableCells"),
                  style: "absolute",
                  value: String(solverChangingRange),
                  onChange: (text: string) => {
                    setSolverChangingRange(text);
                    setSolverPending(null);
                    setSolverStatus(null);
                  },
                })
              }
              testId="spreadsheet-solver-changing-pick"
            />
          </span>
        </label>
        <span className={dialogParts.label}>
          {t("spreadsheet.uiCopy.solverSubjectToConstraints")}
        </span>
        <span className="grid grid-cols-[minmax(0,1fr)_auto] gap-2">
          <ul
            role="listbox"
            aria-label={t("spreadsheet.uiCopy.solverSubjectToConstraints")}
            data-testid="spreadsheet-solver-constraint-list"
            className={`${dialogParts.well} m-0 h-28 list-none overflow-auto p-1`}
          >
            {solverConstraints.map((constraint, index) => (
              <li
                key={index}
                role="option"
                aria-selected={solverSelectedConstraint === index}
                tabIndex={-1}
                data-testid="spreadsheet-solver-constraint-row"
                onClick={() => setSolverSelectedConstraint(index)}
                onDoubleClick={() => openSolverConstraintEditor(index)}
                className="cursor-default rounded px-1 font-mono text-[12px]"
                style={
                  solverSelectedConstraint === index
                    ? { background: "var(--console-accent-strong, #087c73)", color: "var(--primary-foreground, #fff)" }
                    : undefined
                }
              >
                {solverConstraintText(constraint)}
              </li>
            ))}
          </ul>
          <span className="flex flex-col gap-1">
            <button
              type="button"
              className={dialogParts.compactButton}
              onClick={() => openSolverConstraintEditor(null)}
              data-testid="spreadsheet-solver-constraint-add"
            >
              <SpreadsheetDialogIcon name="add" />
              {t("spreadsheet.solverAddConstraint")}
            </button>
            <button
              type="button"
              className={dialogParts.compactButton}
              disabled={solverSelectedConstraint === null}
              onClick={() =>
                solverSelectedConstraint !== null &&
                openSolverConstraintEditor(solverSelectedConstraint)
              }
              data-testid="spreadsheet-solver-constraint-change"
            >
              <SpreadsheetDialogIcon name="edit" />
              {t("spreadsheet.uiCopy.solverChange")}
            </button>
            <button
              type="button"
              className={dialogParts.compactButton}
              disabled={solverSelectedConstraint === null}
              onClick={() => {
                if (solverSelectedConstraint === null) return;
                removeSolverConstraint(solverSelectedConstraint);
                setSolverSelectedConstraint(null);
              }}
              data-testid="spreadsheet-solver-constraint-remove"
            >
              <SpreadsheetDialogIcon name="delete" />
              {t("spreadsheet.uiCopy.solverDelete")}
            </button>
            <button
              type="button"
              className={dialogParts.compactButton}
              onClick={() => setSolverSubDialog({ kind: "reset" })}
              data-testid="spreadsheet-solver-reset"
            >
              <SpreadsheetDialogIcon name="reset" />
              {t("spreadsheet.uiCopy.solverResetAll")}
            </button>
            <button
              type="button"
              className={dialogParts.compactButton}
              onClick={() => setSolverSubDialog({ kind: "model", range: "" })}
              data-testid="spreadsheet-solver-load-save"
            >
              <SpreadsheetDialogIcon name="save" />
              {t("spreadsheet.uiCopy.solverLoadSave")}
            </button>
          </span>
        </span>
        <label className={dialogParts.check}>
          <input
            type="checkbox"
            checked={solverNonNegative}
            onChange={(event) => {
              setSolverNonNegative(event.target.checked);
              setSolverPending(null);
              setSolverStatus(null);
            }}
            data-testid="spreadsheet-solver-non-negative"
          />
          <span>{t("spreadsheet.uiCopy.solverNonNegative")}</span>
        </label>
        <span className="flex items-end gap-2">
          <label className={`${dialogParts.field} flex-1`}>
            <span className={dialogParts.label}>
              {t("spreadsheet.uiCopy.solverSelectMethod")}
            </span>
            <select
              value={solverMethod}
              onChange={(event) => {
                setSolverMethod(event.target.value as SolverDialogMethod);
                setSolverPending(null);
                setSolverStatus(null);
              }}
              data-testid="spreadsheet-solver-method"
            >
              <option value="grg">{t("spreadsheet.uiCopy.solverMethodGrg")}</option>
              <option value="simplex">{t("spreadsheet.uiCopy.solverMethodSimplex")}</option>
              <option value="evolutionary">{t("spreadsheet.uiCopy.solverMethodEvolutionary")}</option>
            </select>
          </label>
          <button
            type="button"
            className={dialogParts.button}
            onClick={() => setSolverSubDialog({ kind: "options" })}
            data-testid="spreadsheet-solver-options"
          >
            <SpreadsheetDialogIcon name="options" />
            {t("spreadsheet.uiCopy.solverOptions")}
          </button>
        </span>
        <p className={`${dialogParts.hint} m-0`} data-testid="spreadsheet-solver-method-help">
          {t("spreadsheet.uiCopy.solverMethodHelp")}
        </p>
        {solverStatus ? (
          <span
            role="status"
            data-testid="spreadsheet-solver-status"
            data-status={solverStatus.kind}
            className={dialogParts.message}
            data-tone={solverStatus.kind === "success" ? "success" : "error"}
          >
            {solverStatus.text}
            {solverTrialOpen ? (
              <button
                type="button"
                className={dialogParts.compactButton}
                style={{ marginInlineStart: 8 }}
                onClick={() => {
                  solverTrialContinueRef.current = true;
                  void runSolver();
                }}
                data-testid="spreadsheet-solver-trial-continue"
              >
                {t("spreadsheet.uiCopy.solverContinue")}
              </button>
            ) : null}
          </span>
        ) : null}
      </OfficeDialogShell>
  );
}
