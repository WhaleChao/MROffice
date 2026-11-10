"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { type SolverDialogOp } from "../../../../lib/spreadsheet/solver-model";
import { type Dispatch, type SetStateAction } from "react";
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import { type SpreadsheetRangePick } from "./useSpreadsheetRangePick";

/** What renderSpreadsheetSolverConstraintDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetSolverConstraintDialogInputs {
  rangePick: SpreadsheetRangePick;
  commitSolverConstraintEditor: (close: boolean) => void;
  setSolverSubDialog: Dispatch<SetStateAction<{ kind: "constraint"; index: number | null; cellRef: string; op: SolverDialogOp; bound: string; } | { kind: "options"; } | { kind: "model"; range: string; } | { kind: "reset"; } | null>>;
  solverConstraintError: string | null;
  solverOpen: boolean;
  solverSubDialog: { kind: "constraint"; index: number | null; cellRef: string; op: SolverDialogOp; bound: string; } | { kind: "options"; } | { kind: "model"; range: string; } | { kind: "reset"; } | null;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetSolverConstraintDialog({
  rangePick,
  commitSolverConstraintEditor,
  setSolverSubDialog,
  solverConstraintError,
  solverOpen,
  solverSubDialog,
  t,
}: RenderSpreadsheetSolverConstraintDialogInputs) {
  // Collapsed while a range of it is picked on the sheet.
  if (rangePick.picking("spreadsheet-solver-constraint")) return null;
  return (
      <OfficeDialogShell
        open={solverOpen && solverSubDialog?.kind === "constraint"}
        title={t(
          solverSubDialog?.kind === "constraint" && solverSubDialog.index !== null
            ? "spreadsheet.uiCopy.solverChangeConstraint"
            : "spreadsheet.uiCopy.solverAddConstraintTitle",
        )}
        description={null}
        testIdPrefix="spreadsheet-solver-constraint-editor"
        defaultActionTestId="spreadsheet-solver-constraint-ok"
        defaultActionLabel={t("spreadsheet.fmtOk")}
        cancelActionLabel={t("spreadsheet.solverCancel")}
        onDefaultAction={() => commitSolverConstraintEditor(true)}
        onCancel={() => setSolverSubDialog(null)}
        footerStart={
          <button
            type="button"
            className={dialogParts.button}
            onClick={() => commitSolverConstraintEditor(false)}
            data-testid="spreadsheet-solver-constraint-add-next"
          >
            {t("spreadsheet.solverAddConstraint")}
          </button>
        }
        className="sm:max-w-[440px]"
      >
        {solverSubDialog?.kind === "constraint" ? (
          <span className="grid grid-cols-[minmax(0,1fr)_auto_minmax(0,1fr)] items-end gap-2">
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverCellReference")}</span>
              <span style={{ display: "flex", gap: 4 }}>
                <input
                  value={solverSubDialog.cellRef}
                  onChange={(event) =>
                    setSolverSubDialog({ ...solverSubDialog, cellRef: event.target.value })
                  }
                  spellCheck={false}
                  data-testid="spreadsheet-solver-constraint-cell"
                  style={{ flex: 1, minWidth: 0 }}
                />
                <SpreadsheetCollapseDialogButton
                  label={t("spreadsheet.cfDialog.collapse")}
                  onClick={() =>
                    rangePick.start({
                      id: "spreadsheet-solver-constraint-cell",
                      dialogLabel: t( solverSubDialog?.kind === "constraint" && solverSubDialog.index !== null ? "spreadsheet.uiCopy.solverChangeConstraint" : "spreadsheet.uiCopy.solverAddConstraintTitle", ),
                      inputLabel: t("spreadsheet.uiCopy.solverCellReference"),
                      style: "absolute",
                      value: String(solverSubDialog.cellRef),
                      onChange: (text: string) => setSolverSubDialog({ ...solverSubDialog, cellRef: text }),
                    })
                  }
                  testId="spreadsheet-solver-constraint-cell-pick"
                />
              </span>
            </label>
            <select
              value={solverSubDialog.op}
              onChange={(event) =>
                setSolverSubDialog({
                  ...solverSubDialog,
                  op: event.target.value as SolverDialogOp,
                  bound:
                    event.target.value === "int" || event.target.value === "bin" || event.target.value === "dif"
                      ? ""
                      : solverSubDialog.bound,
                })
              }
              aria-label={t("spreadsheet.uiCopy.solverRelation")}
              data-testid="spreadsheet-solver-constraint-op"
            >
              <option value="<=">{"<="}</option>
              <option value="==">{"="}</option>
              <option value=">=">{">="}</option>
              <option value="int">int</option>
              <option value="bin">bin</option>
              <option value="dif">dif</option>
            </select>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.uiCopy.solverConstraint")}</span>
              <input
                value={
                  solverSubDialog.op === "int"
                    ? t("spreadsheet.uiCopy.solverInteger")
                    : solverSubDialog.op === "bin"
                      ? t("spreadsheet.uiCopy.solverBinary")
                      : solverSubDialog.op === "dif"
                        ? t("spreadsheet.uiCopy.solverAllDifferent")
                        : solverSubDialog.bound
                }
                readOnly={solverSubDialog.op === "int" || solverSubDialog.op === "bin" || solverSubDialog.op === "dif"}
                onChange={(event) =>
                  setSolverSubDialog({ ...solverSubDialog, bound: event.target.value })
                }
                spellCheck={false}
                data-testid="spreadsheet-solver-constraint-bound"
              />
            </label>
          </span>
        ) : null}
        {solverConstraintError ? (
          <span role="alert" className={dialogParts.error} data-testid="spreadsheet-solver-constraint-error">
            {solverConstraintError}
          </span>
        ) : null}
      </OfficeDialogShell>
  );
}
