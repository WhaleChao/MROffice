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
import { type SheetModel } from "./spreadsheetSheetModel";
import { type Dispatch, type SetStateAction } from "react";
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import { type SpreadsheetRangePick } from "./useSpreadsheetRangePick";

/** What renderSpreadsheetGoalSeekDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetGoalSeekDialogInputs {
  rangePick: SpreadsheetRangePick;
  applyGoalSeekResult: () => void;
  goalSeekChangeCell: string;
  goalSeekOpen: boolean;
  goalSeekPending: { sourceSheet: SheetModel; row: number; col: number; value: string; } | null;
  goalSeekSetCell: string;
  goalSeekSetInputRef: RefObject<HTMLInputElement | null>;
  goalSeekStatus: { kind: "success" | "error"; text: string; } | null;
  goalSeekTarget: string;
  runGoalSeek: () => Promise<void>;
  setGoalSeekChangeCell: Dispatch<SetStateAction<string>>;
  setGoalSeekOpen: Dispatch<SetStateAction<boolean>>;
  setGoalSeekPending: Dispatch<SetStateAction<{ sourceSheet: SheetModel; row: number; col: number; value: string; } | null>>;
  setGoalSeekSetCell: Dispatch<SetStateAction<string>>;
  setGoalSeekStatus: Dispatch<SetStateAction<{ kind: "success" | "error"; text: string; } | null>>;
  setGoalSeekTarget: Dispatch<SetStateAction<string>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  whatIfAnalysisReturnFocusTarget: () => HTMLElement | null;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetGoalSeekDialog({
  rangePick,
  applyGoalSeekResult,
  goalSeekChangeCell,
  goalSeekOpen,
  goalSeekPending,
  goalSeekSetCell,
  goalSeekSetInputRef,
  goalSeekStatus,
  goalSeekTarget,
  runGoalSeek,
  setGoalSeekChangeCell,
  setGoalSeekOpen,
  setGoalSeekPending,
  setGoalSeekSetCell,
  setGoalSeekStatus,
  setGoalSeekTarget,
  t,
  whatIfAnalysisReturnFocusTarget,
}: RenderSpreadsheetGoalSeekDialogInputs) {
  // Collapsed while a range of it is picked on the sheet.
  if (rangePick.picking("spreadsheet-goal-seek")) return null;
  return (
      <OfficeDialogShell
        open={goalSeekOpen}
        title={t("spreadsheet.goalSeekMenu")}
        description={null}
        testIdPrefix="spreadsheet-goal-seek"
        defaultActionTestId="spreadsheet-goal-seek-run"
        defaultActionData={{ "data-mode": goalSeekPending ? "apply" : "solve" }}
        defaultActionLabel={t("spreadsheet.fmtOk")}
        cancelActionLabel={t("spreadsheet.goalSeekCancel")}
        onDefaultAction={goalSeekPending ? applyGoalSeekResult : runGoalSeek}
        onCancel={() => {
          setGoalSeekOpen(false);
          setGoalSeekPending(null);
          setGoalSeekStatus(null);
        }}
        initialFocusRef={goalSeekSetInputRef}
        returnFocusTarget={whatIfAnalysisReturnFocusTarget}
        className="sm:max-w-[360px]"
      >
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.goalSeekSetCellLabel")}
          </span>
          <span style={{ display: "flex", gap: 4 }}>
            <input
              ref={goalSeekSetInputRef}
              value={goalSeekSetCell}
              onChange={(event) => {
                setGoalSeekSetCell(event.target.value);
                setGoalSeekPending(null);
                setGoalSeekStatus(null);
              }}
              placeholder={t("spreadsheet.goalSeekCellPlaceholder")}
              data-testid="spreadsheet-goal-seek-set"
              style={{ flex: 1, minWidth: 0 }}
            />
            <SpreadsheetCollapseDialogButton
              label={t("spreadsheet.cfDialog.collapse")}
              onClick={() =>
                rangePick.start({
                  id: "spreadsheet-goal-seek-set",
                  dialogLabel: t("spreadsheet.goalSeekMenu"),
                  inputLabel: t("spreadsheet.goalSeekSetCellLabel"),
                  style: "absolute",
                  value: String(goalSeekSetCell),
                  onChange: (text: string) => {
                    setGoalSeekSetCell(text);
                    setGoalSeekPending(null);
                    setGoalSeekStatus(null);
                  },
                })
              }
              testId="spreadsheet-goal-seek-set-pick"
            />
          </span>
        </label>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.goalSeekToValueLabel")}
          </span>
          <input
            value={goalSeekTarget}
            onChange={(event) => {
              setGoalSeekTarget(event.target.value);
              setGoalSeekPending(null);
              setGoalSeekStatus(null);
            }}
            placeholder={t("spreadsheet.goalSeekValuePlaceholder")}
            inputMode="decimal"
            data-testid="spreadsheet-goal-seek-target"
          />
        </label>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.goalSeekByChangingLabel")}
          </span>
          <span style={{ display: "flex", gap: 4 }}>
            <input
              value={goalSeekChangeCell}
              onChange={(event) => {
                setGoalSeekChangeCell(event.target.value);
                setGoalSeekPending(null);
                setGoalSeekStatus(null);
              }}
              placeholder={t("spreadsheet.goalSeekCellPlaceholder")}
              data-testid="spreadsheet-goal-seek-change"
              style={{ flex: 1, minWidth: 0 }}
            />
            <SpreadsheetCollapseDialogButton
              label={t("spreadsheet.cfDialog.collapse")}
              onClick={() =>
                rangePick.start({
                  id: "spreadsheet-goal-seek-change",
                  dialogLabel: t("spreadsheet.goalSeekMenu"),
                  inputLabel: t("spreadsheet.goalSeekByChangingLabel"),
                  style: "absolute",
                  value: String(goalSeekChangeCell),
                  onChange: (text: string) => {
                    setGoalSeekChangeCell(text);
                    setGoalSeekPending(null);
                    setGoalSeekStatus(null);
                  },
                })
              }
              testId="spreadsheet-goal-seek-change-pick"
            />
          </span>
        </label>
        {goalSeekStatus ? (
          <span
            role="status"
            data-testid="spreadsheet-goal-seek-status"
            data-status={goalSeekStatus.kind}
            className={dialogParts.message}
            data-tone={goalSeekStatus.kind === "success" ? "success" : "error"}
          >
            {goalSeekStatus.text}
          </span>
        ) : null}
      </OfficeDialogShell>
  );
}
