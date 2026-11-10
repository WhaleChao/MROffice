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
import { normalizeScenarios } from "../../../../lib/spreadsheet/scenarios";
import { type SheetModel } from "./spreadsheetSheetModel";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetScenarioMergeDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetScenarioMergeDialogInputs {
  commitScenarioMerge: () => void;
  scenarioMergeDraft: { sheetId: string; sourceSheetId: string | null; } | null;
  scenarioMergeFocusRef: RefObject<HTMLButtonElement | null>;
  scenarioOpen: boolean;
  setScenarioMergeDraft: Dispatch<SetStateAction<{ sheetId: string; sourceSheetId: string | null; } | null>>;
  sheets: SheetModel[];
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  title: string;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetScenarioMergeDialog({
  commitScenarioMerge,
  scenarioMergeDraft,
  scenarioMergeFocusRef,
  scenarioOpen,
  setScenarioMergeDraft,
  sheets,
  t,
  title,
}: RenderSpreadsheetScenarioMergeDialogInputs) {
  return (
      <OfficeDialogShell
        open={scenarioOpen && scenarioMergeDraft !== null}
        title={t("spreadsheet.scenarioMergeTitle")}
        description={null}
        testIdPrefix="spreadsheet-scenario-merge"
        defaultActionLabel={t("spreadsheet.fmtOk")}
        cancelActionLabel={t("common.cancel")}
        onDefaultAction={commitScenarioMerge}
        onCancel={() => setScenarioMergeDraft(null)}
        initialFocusRef={scenarioMergeFocusRef}
        className="sm:max-w-[360px]"
      >
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>
            {t("spreadsheet.scenarioMergeBook")}
          </span>
          <select
            value="current"
            onChange={() => undefined}
            data-testid="spreadsheet-scenario-merge-book"
          >
            <option value="current">{title}</option>
          </select>
        </label>
        <span className={dialogParts.label}>
          {t("spreadsheet.scenarioMergeSheet")}
        </span>
        <span
          role="listbox"
          aria-label={t("spreadsheet.scenarioMergeSheet")}
          data-testid="spreadsheet-scenario-merge-sheets"
          className={dialogParts.list}
          style={{ maxHeight: 160, minHeight: 96, overflowY: "auto" }}
        >
          {sheets
            .filter((candidate) => !candidate.chartSheet)
            .map((candidate, index, list) => {
              const selected = candidate.id === scenarioMergeDraft?.sourceSheetId;
              return (
                <button
                  key={candidate.id}
                  ref={selected ? scenarioMergeFocusRef : undefined}
                  type="button"
                  role="option"
                  aria-selected={selected}
                  tabIndex={selected ? 0 : -1}
                  onClick={() =>
                    setScenarioMergeDraft((current) =>
                      current ? { ...current, sourceSheetId: candidate.id } : current,
                    )
                  }
                  onKeyDown={(event) => {
                    const step =
                      event.key === "ArrowDown" ? 1 : event.key === "ArrowUp" ? -1 : 0;
                    if (step === 0) return;
                    event.preventDefault();
                    const next = list[index + step];
                    if (!next) return;
                    setScenarioMergeDraft((current) =>
                      current ? { ...current, sourceSheetId: next.id } : current,
                    );
                    const sibling =
                      step > 0
                        ? event.currentTarget.nextElementSibling
                        : event.currentTarget.previousElementSibling;
                    if (sibling instanceof HTMLElement) sibling.focus();
                  }}
                  data-testid="spreadsheet-scenario-merge-sheet"
                  data-sheet-id={candidate.id}
                  className={dialogParts.listItem}
                >
                  {candidate.name}
                </button>
              );
            })}
        </span>
        <span
          role="status"
          data-testid="spreadsheet-scenario-merge-count"
          className={dialogParts.hint}
        >
          {t("spreadsheet.scenarioMergeCount", {
            count: normalizeScenarios(
              sheets.find(
                (candidate) => candidate.id === scenarioMergeDraft?.sourceSheetId,
              )?.scenarios,
            ).length,
          })}
        </span>
      </OfficeDialogShell>
  );
}
