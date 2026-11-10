"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type BackgroundCalculationClient } from "@mrnote/core/web/lib/spreadsheet/background-calculation-client.ts";
import { formatGoalSeekValue } from "../../../../lib/spreadsheet/goal-seek";
import type { SolverPlanResult } from "@/lib/spreadsheet/what-if-buttons";
import { useCallback, type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { parseCellRef, type SheetEvaluator, type EngineValue, type WorkbookSheetInput } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { type SolverDialogConstraint, type SolverDialogExtraOptions, type SolverDialogMethod } from "../../../../lib/spreadsheet/solver-model";
import { type PivotRange, type PivotMatrix } from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import { type PivotChartReport } from "@mrnote/core/web/lib/spreadsheet/pivot-output.ts";
import { type SheetModel } from "./spreadsheetSheetModel";
import { columnLabel } from "./spreadsheetRangeA1";
import { parseSolverCells } from "./spreadsheetDialogDrafts";
import { type Workbook } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { type PivotOutputPlan } from "@mrnote/core/web/lib/spreadsheet/pivot-output.ts";
import { type Dispatch, type SetStateAction } from "react";

/** What useSpreadsheetRunSolver reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetRunSolverInputs {
  formulaWorkbook: { evaluator: SheetEvaluator; evaluatorFor: (sheetName: string) => SheetEvaluator; workbook: Workbook; memoSnapshot: () => ReadonlyMap<string, EngineValue>; pivotPlans: Map<string, PivotOutputPlan>; pivotNativeFootprints: Map<string, PivotRange[]>; pivotChartReports: Map<string, PivotChartReport>; pivotChartMatrices: Map<string, PivotMatrix>; inputs: WorkbookSheetInput[]; pivotSheetIds: string[]; };
  isEditableCell: (row: number, col: number) => boolean;
  sessionWhatIf: (input: Parameters<BackgroundCalculationClient["whatIf"]>[0]) => Promise<unknown | null>;
  setSolverPending: Dispatch<SetStateAction<{ sourceSheet: SheetModel; values: readonly { row: number; col: number; value: string; }[]; } | null>>;
  setSolverStatus: Dispatch<SetStateAction<{ kind: "success" | "error"; text: string; } | null>>;
  setSolverTrialOpen: Dispatch<SetStateAction<boolean>>;
  sheet: SheetModel;
  showProtectNotice: (message?: string) => void;
  solverChangingRange: string;
  solverConstraints: SolverDialogConstraint[];
  solverExtra: SolverDialogExtraOptions;
  solverGoalKind: "min" | "max" | "value";
  solverGoalValue: string;
  solverIgnoreIntegers: boolean;
  solverIterations: string;
  solverMethod: SolverDialogMethod;
  solverNonNegative: boolean;
  solverObjectiveCell: string;
  solverPrecision: string;
  solverTrialContinueRef: RefObject<boolean>;
  solverTrialRef: RefObject<{ budget: number; } | null>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function useSpreadsheetRunSolver({
  formulaWorkbook,
  isEditableCell,
  sessionWhatIf,
  setSolverPending,
  setSolverStatus,
  setSolverTrialOpen,
  sheet,
  showProtectNotice,
  solverChangingRange,
  solverConstraints,
  solverExtra,
  solverGoalKind,
  solverGoalValue,
  solverIgnoreIntegers,
  solverIterations,
  solverMethod,
  solverNonNegative,
  solverObjectiveCell,
  solverPrecision,
  solverTrialContinueRef,
  solverTrialRef,
  t,
}: SpreadsheetRunSolverInputs) {
  // Solver (Excel/WPS "Data > Solver"): find values for a set of CHANGING cells
  // that minimize / maximize / hit a target value of an OBJECTIVE formula cell,
  // subject to optional simple constraint cells. The numeric optimizer lives in
  // solver.ts (pure + injectable). Here we validate the inputs, run the shared
  // plan (one probe of the sheet per candidate vector; in the calculation
  // session in server mode), and hold a successful
  // vector for the explicit Solver Results acceptance step. Only that step writes
  // through replaceActiveSheet(updateCellValue(...)); protection is checked both
  // before exposing and immediately before committing a candidate.
  const runSolver = useCallback(async () => {
    setSolverPending(null);
    const objectiveRef = parseCellRef(solverObjectiveCell);
    if (!objectiveRef) {
      setSolverStatus({
        kind: "error",
        text: t("spreadsheet.solverBadObjective"),
      });
      return;
    }
    // The objective should be a formula, else nothing depends on the inputs.
    const objectiveCell =
      sheet.cells[`${objectiveRef.row}:${objectiveRef.col}`];
    if (
      !objectiveCell?.formula &&
      !String(objectiveCell?.value ?? "")
        .trim()
        .startsWith("=")
    ) {
      setSolverStatus({
        kind: "error",
        text: t("spreadsheet.solverObjectiveNotFormula"),
      });
      return;
    }
    const changingCells = parseSolverCells(solverChangingRange);
    if (!changingCells || changingCells.length === 0) {
      setSolverStatus({
        kind: "error",
        text: t("spreadsheet.solverBadChanging"),
      });
      return;
    }
    // None of the changing cells may be a formula (no single free variable to
    // vary) — mirrors Goal Seek's constant-input guard, per cell.
    for (const cell of changingCells) {
      const c = sheet.cells[`${cell.row}:${cell.col}`];
      if (
        c?.formula ||
        String(c?.value ?? "")
          .trim()
          .startsWith("=")
      ) {
        setSolverStatus({
          kind: "error",
          text: t("spreadsheet.solverChangingIsFormula"),
        });
        return;
      }
    }
    // For a "value" goal, the target value field must be numeric.
    let goalTarget: number | undefined;
    if (solverGoalKind === "value") {
      const parsed = Number(solverGoalValue.trim());
      if (solverGoalValue.trim() === "" || !Number.isFinite(parsed)) {
        setSolverStatus({
          kind: "error",
          text: t("spreadsheet.solverBadGoalValue"),
        });
        return;
      }
      goalTarget = parsed;
    }

    // Parse + validate every constraint row up front (XLS-A073): a relation
    // holds each cell of its Cell Reference against a number or a cell's
    // value; int / bin / dif name variable cells only (Support: they "apply
    // only to decision variable cells"). Blank rows are ignored.
    const variableIndex = new Map(
      changingCells.map((cell, index) => [`${cell.row}:${cell.col}`, index] as const),
    );
    const parsedConstraints: {
      ref: { row: number; col: number };
      op: "<=" | ">=" | "==";
      bound: number;
    }[] = [];
    const integers: number[] = [];
    const binaries: number[] = [];
    const allDifferent: number[][] = [];
    for (const c of solverConstraints) {
      const refText = c.cellRef.trim();
      if (refText === "") continue;
      const cells = parseSolverCells(refText);
      if (!cells) {
        setSolverStatus({
          kind: "error",
          text: t("spreadsheet.solverBadConstraintCell"),
        });
        return;
      }
      if (c.op === "int" || c.op === "bin" || c.op === "dif") {
        const indices = cells.map((cell) => variableIndex.get(`${cell.row}:${cell.col}`));
        if (indices.some((index) => index === undefined)) {
          setSolverStatus({ kind: "error", text: t("spreadsheet.uiCopy.solverWholeNeedsVariables") });
          return;
        }
        // Options > Ignore Integer Constraints.
        if (solverIgnoreIntegers) continue;
        const list = indices as number[];
        if (c.op === "int") integers.push(...list);
        else if (c.op === "bin") binaries.push(...list);
        else allDifferent.push(list);
        continue;
      }
      const boundText = c.bound.trim();
      const boundRef = Number.isFinite(Number(boundText)) ? null : parseCellRef(boundText);
      const boundValue = boundRef
        ? formulaWorkbook.evaluator.getComputed(boundRef.row, boundRef.col)
        : Number(boundText);
      const bound = typeof boundValue === "number" ? boundValue : Number(boundValue);
      if (boundText === "" || !Number.isFinite(bound)) {
        setSolverStatus({
          kind: "error",
          text: t("spreadsheet.solverBadConstraintBound"),
        });
        return;
      }
      for (const ref of cells) parsedConstraints.push({ ref, op: c.op, bound });
    }
    const precision = Number(solverPrecision);
    const iterations = Number(solverIterations);
    if (!(precision > 0 && precision < 1) || !(Number.isInteger(iterations) && iterations > 0)) {
      setSolverStatus({ kind: "error", text: t("spreadsheet.uiCopy.solverBadOptions") });
      return;
    }
    // XLS-A073 Solver Options: a blank box is "no limit" / Excel's default;
    // anything else must be a number in the option's range.
    const optionNumber = (text: string, ok: (value: number) => boolean): number | undefined | null => {
      const trimmed = text.trim();
      if (trimmed === "") return undefined;
      const value = Number(trimmed);
      return Number.isFinite(value) && ok(value) ? value : null;
    };
    const extraValues = {
      maxTimeSeconds: optionNumber(solverExtra.maxTime, (v) => v > 0 && v <= 32_767),
      integerOptimality: optionNumber(solverExtra.integerOptimality, (v) => v >= 0 && v <= 100),
      maxSubproblems: optionNumber(solverExtra.maxSubproblems, (v) => Number.isInteger(v) && v > 0),
      maxFeasibleSolutions: optionNumber(solverExtra.maxFeasibleSolutions, (v) => Number.isInteger(v) && v > 0),
      convergence: optionNumber(solverExtra.convergence, (v) => v > 0 && v < 1),
      populationSize: optionNumber(solverExtra.populationSize, (v) => Number.isInteger(v) && v >= 0 && v <= 1_000),
      randomSeed: optionNumber(solverExtra.randomSeed, (v) => Number.isInteger(v) && v >= 0),
      mutationRate: optionNumber(solverExtra.mutationRate, (v) => v > 0 && v < 1),
      maxTimeNoImprovementSeconds: optionNumber(solverExtra.maxTimeNoImprovement, (v) => v > 0 && v <= 32_767),
    };
    if (Object.values(extraValues).some((value) => value === null)) {
      setSolverStatus({ kind: "error", text: t("spreadsheet.uiCopy.solverBadOptions") });
      return;
    }
    // Show Iteration Results: the search pauses after each sweep over the
    // variables (Excel's Show Trial Solution); Continue widens the budget.
    const sweep = 2 * (changingCells.length + 1) * (parsedConstraints.length + 1);
    const fullBudget = Math.min(100_000, iterations * sweep);
    const continuing = solverTrialContinueRef.current;
    solverTrialContinueRef.current = false;
    const trialBudget = solverExtra.showIterations
      ? Math.min(fullBudget, continuing && solverTrialRef.current ? solverTrialRef.current.budget + sweep : sweep)
      : null;
    solverTrialRef.current = trialBudget === null ? null : { budget: trialBudget };
    setSolverTrialOpen(false);

    // The search is the shared plan (what-if-buttons solverPlanAsync): one
    // probe of the sheet per candidate, in the calculation session (never here).
    const model = {
      objective: objectiveRef,
      changing: changingCells,
      goal: solverGoalKind === "value" ? { kind: "value" as const, target: goalTarget } : { kind: solverGoalKind },
      constraints: parsedConstraints.map((c) => ({ cell: c.ref, op: c.op, bound: c.bound })),
      ...(integers.length ? { integers } : {}),
      ...(binaries.length ? { binaries } : {}),
      ...(allDifferent.length ? { allDifferent } : {}),
      nonNegative: solverNonNegative,
      method: solverMethod,
      precision,
      // Excel's Iterations, as search sweeps over every variable and
      // constraint (the default 100 is the engine's own default budget).
      maxIterations: trialBudget ?? fullBudget,
      // A Population Size of 0 is Excel's "choose for me": left out.
      ...Object.fromEntries(Object.entries(extraValues).filter(([key, value]) => typeof value === "number" && !(key === "populationSize" && value === 0))),
      ...(extraValues.integerOptimality !== undefined ? { integerOptimality: (extraValues.integerOptimality as number) / 100 } : {}),
      scaling: solverExtra.scaling,
      multistart: solverExtra.multistart,
      requireBounds: solverExtra.requireBounds,
    };
    const result = await sessionWhatIf({ kind: "solver", sheetId: sheet.id, ...model }) as SolverPlanResult | null;
    if (!result) {
      setSolverStatus({ kind: "error", text: t("spreadsheet.calculationStale.command") });
      return;
    }

    // Show Iteration Results / Max Time / Iterations: the best point so far is
    // a trial solution the user may keep (Excel's Show Trial Solution and
    // "maximum ... limit was reached; continue anyway?").
    const stoppedAt = (result as { reason?: string }).reason;
    if (!result.solved && (stoppedAt === "iteration-limit" || stoppedAt === "time-limit") && result.x.length === changingCells.length) {
      const trial = Boolean(trialBudget !== null && stoppedAt === "iteration-limit" && trialBudget < fullBudget);
      setSolverTrialOpen(trial);
      setSolverPending({
        sourceSheet: sheet,
        values: changingCells.map((cell, index) => ({ row: cell.row, col: cell.col, value: formatGoalSeekValue(result.x[index]) })),
      });
      setSolverStatus({
        kind: "error",
        text: trial
          ? t("spreadsheet.uiCopy.solverTrialSolution", {
              cell: `${columnLabel(objectiveRef.col)}${objectiveRef.row + 1}`,
              value: formatGoalSeekValue(result.objective ?? 0),
            })
          : t(stoppedAt === "time-limit" ? "spreadsheet.uiCopy.solverTimeLimit" : "spreadsheet.uiCopy.solverIterationLimit"),
      });
      return;
    }

    if (!result.solved) {
      // No write happened, so there is nothing to restore — just report.
      // A method that refused the model says why (Excel's own messages).
      const reasons = {
        nonlinear: "spreadsheet.uiCopy.solverNotLinear",
        "needs-bounds": "spreadsheet.uiCopy.solverNeedsBounds",
        unbounded: "spreadsheet.uiCopy.solverNotConverge",
        infeasible: "spreadsheet.uiCopy.solverInfeasible",
        "time-limit": "spreadsheet.uiCopy.solverTimeLimit",
        "iteration-limit": "spreadsheet.uiCopy.solverIterationLimit",
      } as const;
      setSolverStatus({
        kind: "error",
        text: result.reason
          ? t(reasons[result.reason])
          : t("spreadsheet.solverNoSolution", {
              cell: `${columnLabel(objectiveRef.col)}${objectiveRef.row + 1}`,
            }),
      });
      return;
    }

    // PROTECTION GUARD (whole operation): refuse the write if ANY changing cell
    // is locked on a protected sheet. updateCellValue does NOT enforce this, so
    // we check every target before touching the grid — nothing is written if one
    // cell is locked (mirrors the paste / Scenario Manager guard).
    if (changingCells.some((cell) => !isEditableCell(cell.row, cell.col))) {
      showProtectNotice();
      setSolverStatus({ kind: "error", text: t("spreadsheet.solverLocked") });
      return;
    }

    setSolverPending({
      sourceSheet: sheet,
      values: changingCells.map((cell, index) => ({
        row: cell.row,
        col: cell.col,
        value: formatGoalSeekValue(result.x[index]),
      })),
    });
    setSolverStatus({
      kind: "success",
      text: t("spreadsheet.solverSolved", {
        cell: `${columnLabel(objectiveRef.col)}${objectiveRef.row + 1}`,
        value: formatGoalSeekValue(result.objective ?? 0),
        count: changingCells.length,
      }),
    });
  }, [
    solverObjectiveCell,
    solverGoalKind,
    solverGoalValue,
    solverChangingRange,
    solverConstraints,
    solverIgnoreIntegers,
    solverIterations,
    solverMethod,
    solverNonNegative,
    solverPrecision,
    solverExtra,
    formulaWorkbook,
    sessionWhatIf,
    sheet,
    isEditableCell,
    showProtectNotice,
    t,
  ]);
  return {
    runSolver,
  };
}
