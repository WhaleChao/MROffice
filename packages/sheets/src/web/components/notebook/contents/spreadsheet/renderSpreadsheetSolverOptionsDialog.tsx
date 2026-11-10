"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { type SolverDialogExtraOptions, type SolverDialogOp } from "../../../../lib/spreadsheet/solver-model";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetSolverOptionsDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetSolverOptionsDialogInputs {
  setSolverExtra: Dispatch<SetStateAction<SolverDialogExtraOptions>>;
  setSolverIgnoreIntegers: Dispatch<SetStateAction<boolean>>;
  setSolverIterations: Dispatch<SetStateAction<string>>;
  setSolverOptionsTab: Dispatch<SetStateAction<"all" | "grg" | "evolutionary">>;
  setSolverPrecision: Dispatch<SetStateAction<string>>;
  setSolverSubDialog: Dispatch<SetStateAction<{ kind: "constraint"; index: number | null; cellRef: string; op: SolverDialogOp; bound: string; } | { kind: "options"; } | { kind: "model"; range: string; } | { kind: "reset"; } | null>>;
  solverExtra: SolverDialogExtraOptions;
  solverIgnoreIntegers: boolean;
  solverIterations: string;
  solverOpen: boolean;
  solverOptionsTab: "all" | "grg" | "evolutionary";
  solverPrecision: string;
  solverSubDialog: { kind: "constraint"; index: number | null; cellRef: string; op: SolverDialogOp; bound: string; } | { kind: "options"; } | { kind: "model"; range: string; } | { kind: "reset"; } | null;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetSolverOptionsDialog({
  setSolverExtra,
  setSolverIgnoreIntegers,
  setSolverIterations,
  setSolverOptionsTab,
  setSolverPrecision,
  setSolverSubDialog,
  solverExtra,
  solverIgnoreIntegers,
  solverIterations,
  solverOpen,
  solverOptionsTab,
  solverPrecision,
  solverSubDialog,
  t,
}: RenderSpreadsheetSolverOptionsDialogInputs) {
  return (
      <OfficeDialogShell
        open={solverOpen && solverSubDialog?.kind === "options"}
        title={t("spreadsheet.uiCopy.solverOptions")}
        description={null}
        testIdPrefix="spreadsheet-solver-options-dialog"
        defaultActionTestId="spreadsheet-solver-options-ok"
        defaultActionLabel={t("spreadsheet.fmtOk")}
        cancelActionLabel={t("spreadsheet.solverCancel")}
        onDefaultAction={() => setSolverSubDialog(null)}
        onCancel={() => setSolverSubDialog(null)}
        className="sm:max-w-[420px]"
        tabStrip={
          <span role="tablist" className={dialogParts.tabs}>
            {(["all", "grg", "evolutionary"] as const).map((tab) => (
              <button
                key={tab}
                type="button"
                role="tab"
                aria-selected={solverOptionsTab === tab}
                className={dialogParts.tab}
                onClick={() => setSolverOptionsTab(tab)}
                data-testid={`spreadsheet-solver-options-tab-${tab}`}
              >
                {t(tab === "all" ? "spreadsheet.uiCopy.solverAllMethods" : tab === "grg" ? "spreadsheet.uiCopy.solverMethodGrg" : "spreadsheet.uiCopy.solverMethodEvolutionary")}
              </button>
            ))}
          </span>
        }
      >
        {solverOptionsTab === "grg" ? (
          <>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverConvergence")}</span>
              <input value={solverExtra.convergence} inputMode="decimal" onChange={(event) => setSolverExtra({ ...solverExtra, convergence: event.target.value })} data-testid="spreadsheet-solver-convergence" />
            </label>
            <fieldset className={dialogParts.field}>
              <legend className={dialogParts.label}>{t("spreadsheet.uiCopy.solverDerivatives")}</legend>
              {(["forward", "central"] as const).map((value) => (
                <label key={value} className={dialogParts.check}>
                  <input type="radio" checked={solverExtra.derivatives === value} onChange={() => setSolverExtra({ ...solverExtra, derivatives: value })} data-testid={`spreadsheet-solver-derivatives-${value}`} />
                  <span>{t(value === "forward" ? "spreadsheet.uiCopy.solverForward" : "spreadsheet.uiCopy.solverCentral")}</span>
                </label>
              ))}
              <p className={dialogParts.hint}>{t("spreadsheet.uiCopy.solverDerivativesNote")}</p>
            </fieldset>
            <label className={dialogParts.check}>
              <input type="checkbox" checked={solverExtra.multistart} onChange={(event) => setSolverExtra({ ...solverExtra, multistart: event.target.checked })} data-testid="spreadsheet-solver-multistart" />
              <span>{t("spreadsheet.uiCopy.solverMultistart")}</span>
            </label>
          </>
        ) : null}
        {solverOptionsTab === "evolutionary" ? (
          <>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverConvergence")}</span>
              <input value={solverExtra.convergence} inputMode="decimal" onChange={(event) => setSolverExtra({ ...solverExtra, convergence: event.target.value })} data-testid="spreadsheet-solver-evo-convergence" />
            </label>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverMutationRate")}</span>
              <input value={solverExtra.mutationRate} inputMode="decimal" onChange={(event) => setSolverExtra({ ...solverExtra, mutationRate: event.target.value })} data-testid="spreadsheet-solver-mutation-rate" />
            </label>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverMaxTimeNoImprovement")}</span>
              <input value={solverExtra.maxTimeNoImprovement} inputMode="numeric" onChange={(event) => setSolverExtra({ ...solverExtra, maxTimeNoImprovement: event.target.value })} data-testid="spreadsheet-solver-max-time-no-improvement" />
            </label>
          </>
        ) : null}
        {solverOptionsTab !== "all" ? (
          <>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverPopulationSize")}</span>
              <input value={solverExtra.populationSize} inputMode="numeric" onChange={(event) => setSolverExtra({ ...solverExtra, populationSize: event.target.value })} data-testid="spreadsheet-solver-population-size" />
            </label>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverRandomSeed")}</span>
              <input value={solverExtra.randomSeed} inputMode="numeric" onChange={(event) => setSolverExtra({ ...solverExtra, randomSeed: event.target.value })} data-testid="spreadsheet-solver-random-seed" />
            </label>
            <label className={dialogParts.check}>
              <input type="checkbox" checked={solverExtra.requireBounds} onChange={(event) => setSolverExtra({ ...solverExtra, requireBounds: event.target.checked })} data-testid="spreadsheet-solver-require-bounds" />
              <span>{t("spreadsheet.uiCopy.solverRequireBounds")}</span>
            </label>
          </>
        ) : null}
        {solverOptionsTab === "all" ? (<>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverConstraintPrecision")}</span>
          <input
            value={solverPrecision}
            onChange={(event) => setSolverPrecision(event.target.value)}
            inputMode="decimal"
            data-testid="spreadsheet-solver-precision"
          />
        </label>
        <label className={dialogParts.check}>
          <input
            type="checkbox"
            checked={solverIgnoreIntegers}
            onChange={(event) => setSolverIgnoreIntegers(event.target.checked)}
            data-testid="spreadsheet-solver-ignore-integers"
          />
          <span>{t("spreadsheet.uiCopy.solverIgnoreIntegers")}</span>
        </label>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverIterations")}</span>
          <input
            value={solverIterations}
            onChange={(event) => setSolverIterations(event.target.value)}
            inputMode="numeric"
            data-testid="spreadsheet-solver-iterations"
          />
        </label>
        <label className={dialogParts.check}>
          <input type="checkbox" checked={solverExtra.scaling} onChange={(event) => setSolverExtra({ ...solverExtra, scaling: event.target.checked })} data-testid="spreadsheet-solver-scaling" />
          <span>{t("spreadsheet.uiCopy.solverUseScaling")}</span>
        </label>
        <label className={dialogParts.check}>
          <input type="checkbox" checked={solverExtra.showIterations} onChange={(event) => setSolverExtra({ ...solverExtra, showIterations: event.target.checked })} data-testid="spreadsheet-solver-show-iterations" />
          <span>{t("spreadsheet.uiCopy.solverShowIterations")}</span>
        </label>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverIntegerOptimality")}</span>
          <input value={solverExtra.integerOptimality} inputMode="decimal" onChange={(event) => setSolverExtra({ ...solverExtra, integerOptimality: event.target.value })} data-testid="spreadsheet-solver-integer-optimality" />
        </label>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverMaxTime")}</span>
          <input value={solverExtra.maxTime} inputMode="numeric" onChange={(event) => setSolverExtra({ ...solverExtra, maxTime: event.target.value })} data-testid="spreadsheet-solver-max-time" />
        </label>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverMaxSubproblems")}</span>
          <input value={solverExtra.maxSubproblems} inputMode="numeric" onChange={(event) => setSolverExtra({ ...solverExtra, maxSubproblems: event.target.value })} data-testid="spreadsheet-solver-max-subproblems" />
        </label>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverMaxFeasible")}</span>
          <input value={solverExtra.maxFeasibleSolutions} inputMode="numeric" onChange={(event) => setSolverExtra({ ...solverExtra, maxFeasibleSolutions: event.target.value })} data-testid="spreadsheet-solver-max-feasible" />
        </label>
        </>) : null}
      </OfficeDialogShell>
  );
}
