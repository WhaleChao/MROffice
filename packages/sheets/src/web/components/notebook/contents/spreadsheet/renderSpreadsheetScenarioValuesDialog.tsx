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
import { parseCellRef } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { columnLabel } from "./spreadsheetRangeA1";
import { type ScenarioFormDraft } from "./spreadsheetWindowLocalTypes";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetScenarioValuesDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetScenarioValuesDialogInputs {
  commitScenarioValues: () => void;
  scenarioFirstValueInputRef: RefObject<HTMLInputElement | null>;
  scenarioOpen: boolean;
  scenarioValuesDraft: { form: ScenarioFormDraft; refs: readonly string[]; values: readonly string[]; } | null;
  setScenarioForm: Dispatch<SetStateAction<ScenarioFormDraft | null>>;
  setScenarioValuesDraft: Dispatch<SetStateAction<{ form: ScenarioFormDraft; refs: readonly string[]; values: readonly string[]; } | null>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetScenarioValuesDialog({
  commitScenarioValues,
  scenarioFirstValueInputRef,
  scenarioOpen,
  scenarioValuesDraft,
  setScenarioForm,
  setScenarioValuesDraft,
  t,
}: RenderSpreadsheetScenarioValuesDialogInputs) {
  return (
      <OfficeDialogShell
        open={scenarioOpen && scenarioValuesDraft !== null}
        title={t("spreadsheet.scenarioValuesTitle")}
        description={t("spreadsheet.scenarioValuesHint")}
        testIdPrefix="spreadsheet-scenario-values"
        defaultActionLabel={t("spreadsheet.fmtOk")}
        cancelActionLabel={t("common.cancel")}
        onDefaultAction={commitScenarioValues}
        onCancel={() => {
          setScenarioValuesDraft(null);
          setScenarioForm(null);
        }}
        initialFocusRef={scenarioFirstValueInputRef}
        className="sm:max-w-[360px]"
      >
        <span
          style={{
            display: "grid",
            gap: 6,
            maxHeight: 220,
            overflowY: "auto",
          }}
        >
          {(scenarioValuesDraft?.refs ?? []).map((ref, index) => {
            const parsed = parseCellRef(ref);
            const label = parsed
              ? `$${columnLabel(parsed.col)}$${parsed.row + 1}`
              : ref;
            return (
              <label
                key={ref}
                style={{
                  display: "grid",
                  gridTemplateColumns: "auto minmax(0, 1fr)",
                  gap: 8,
                  alignItems: "center",
                }}
              >
                <span className={dialogParts.label}>
                  {`${index + 1}: ${label}`}
                </span>
                <input
                  ref={index === 0 ? scenarioFirstValueInputRef : undefined}
                  value={scenarioValuesDraft?.values[index] ?? ""}
                  onChange={(event) => {
                    const value = event.target.value;
                    setScenarioValuesDraft((current) => {
                      if (!current) return current;
                      const values = current.values.slice();
                      values[index] = value;
                      return { ...current, values };
                    });
                  }}
                  data-testid={`spreadsheet-scenario-value-${index}`}
                />
              </label>
            );
          })}
        </span>
      </OfficeDialogShell>
  );
}
