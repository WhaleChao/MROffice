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

/** What renderSpreadsheetDataTableWhatIfDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetDataTableWhatIfDialogInputs {
  rangePick: SpreadsheetRangePick;
  gridActiveCellFocusTarget: () => HTMLElement | null;
  runWhatIfTable: () => Promise<void>;
  setWhatIfInputCell: Dispatch<SetStateAction<string>>;
  setWhatIfOpen: Dispatch<SetStateAction<boolean>>;
  setWhatIfRowInputCell: Dispatch<SetStateAction<string>>;
  setWhatIfStatus: Dispatch<SetStateAction<{ kind: "success" | "error"; text: string; } | null>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  whatIfAnalysisReturnFocusTarget: () => HTMLElement | null;
  whatIfFormulaInputRef: RefObject<HTMLInputElement | null>;
  whatIfInputCell: string;
  whatIfOpen: boolean;
  whatIfReturnToGridRef: RefObject<boolean>;
  whatIfRowInputCell: string;
  whatIfStatus: { kind: "success" | "error"; text: string; } | null;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetDataTableWhatIfDialog({
  rangePick,
  gridActiveCellFocusTarget,
  runWhatIfTable,
  setWhatIfInputCell,
  setWhatIfOpen,
  setWhatIfRowInputCell,
  setWhatIfStatus,
  t,
  whatIfAnalysisReturnFocusTarget,
  whatIfFormulaInputRef,
  whatIfInputCell,
  whatIfOpen,
  whatIfReturnToGridRef,
  whatIfRowInputCell,
  whatIfStatus,
}: RenderSpreadsheetDataTableWhatIfDialogInputs) {
  // Collapsed while a range of it is picked on the sheet.
  if (rangePick.picking("spreadsheet-what-if")) return null;
  return (
      <OfficeDialogShell
        open={whatIfOpen}
        title={t("editorCommands.spreadsheet.data.what_if_table")}
        description={null}
        testIdPrefix="spreadsheet-what-if"
        defaultActionTestId="spreadsheet-what-if-run"
        defaultActionLabel={t("spreadsheet.fmtOk")}
        cancelActionLabel={t("spreadsheet.whatIfCancel")}
        onDefaultAction={runWhatIfTable}
        onCancel={() => {
          setWhatIfOpen(false);
          setWhatIfStatus(null);
        }}
        initialFocusRef={whatIfFormulaInputRef}
        returnFocusTarget={() => {
          if (whatIfReturnToGridRef.current) {
            whatIfReturnToGridRef.current = false;
            return gridActiveCellFocusTarget();
          }
          return whatIfAnalysisReturnFocusTarget();
        }}
        className="sm:max-w-[360px]"
      >
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.uiCopy.dataTableRowInputCell")}
          </span>
          <span style={{ display: "flex", gap: 4 }}>
            <input
              ref={whatIfFormulaInputRef}
              value={whatIfRowInputCell}
              onChange={(event) => {
                setWhatIfRowInputCell(event.target.value);
                if (whatIfStatus) setWhatIfStatus(null);
              }}
              spellCheck={false}
              data-testid="spreadsheet-what-if-row-input"
              style={{ flex: 1, minWidth: 0 }}
            />
            <SpreadsheetCollapseDialogButton
              label={t("spreadsheet.cfDialog.collapse")}
              onClick={() =>
                rangePick.start({
                  id: "spreadsheet-what-if-row-input",
                  dialogLabel: t("editorCommands.spreadsheet.data.what_if_table"),
                  inputLabel: t("spreadsheet.uiCopy.dataTableRowInputCell"),
                  style: "absolute",
                  value: String(whatIfRowInputCell),
                  onChange: (text: string) => {
                    setWhatIfRowInputCell(text);
                    if (whatIfStatus) setWhatIfStatus(null);
                  },
                })
              }
              testId="spreadsheet-what-if-row-input-pick"
            />
          </span>
        </label>
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.uiCopy.dataTableColumnInputCell")}
          </span>
          <span style={{ display: "flex", gap: 4 }}>
            <input
              value={whatIfInputCell}
              onChange={(event) => {
                setWhatIfInputCell(event.target.value);
                if (whatIfStatus) setWhatIfStatus(null);
              }}
              spellCheck={false}
              data-testid="spreadsheet-what-if-input"
              style={{ flex: 1, minWidth: 0 }}
            />
            <SpreadsheetCollapseDialogButton
              label={t("spreadsheet.cfDialog.collapse")}
              onClick={() =>
                rangePick.start({
                  id: "spreadsheet-what-if-input",
                  dialogLabel: t("editorCommands.spreadsheet.data.what_if_table"),
                  inputLabel: t("spreadsheet.uiCopy.dataTableColumnInputCell"),
                  style: "absolute",
                  value: String(whatIfInputCell),
                  onChange: (text: string) => {
                    setWhatIfInputCell(text);
                    if (whatIfStatus) setWhatIfStatus(null);
                  },
                })
              }
              testId="spreadsheet-what-if-input-pick"
            />
          </span>
        </label>
        {whatIfStatus ? (
          <span
            role="status"
            data-testid="spreadsheet-what-if-status"
            data-status={whatIfStatus.kind}
            className={dialogParts.message}
            data-tone={whatIfStatus.kind === "success" ? "success" : "error"}
          >
            {whatIfStatus.text}
          </span>
        ) : null}
      </OfficeDialogShell>
  );
}
