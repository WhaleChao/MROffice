"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetAddCustomViewDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetAddCustomViewDialogInputs {
  addCustomView: () => void;
  commandDialogReturnFocusTarget: () => HTMLElement | null;
  customViewsDialog: { selected: string | null; adding: { name: string; printSettings: boolean; hiddenAndFilters: boolean; } | null; confirmDelete: string | null; } | null;
  setCustomViewsDialog: Dispatch<SetStateAction<{ selected: string | null; adding: { name: string; printSettings: boolean; hiddenAndFilters: boolean; } | null; confirmDelete: string | null; } | null>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetAddCustomViewDialog({
  addCustomView,
  commandDialogReturnFocusTarget,
  customViewsDialog,
  setCustomViewsDialog,
  t,
}: RenderSpreadsheetAddCustomViewDialogInputs) {
  return (
       customViewsDialog?.adding ? (
        // XLS-A030: Add View — Name and "Include in view".
        <OfficeDialogShell
          open
          testIdPrefix="spreadsheet-custom-views-add"
          title={t("spreadsheet.customViews.addTitle")}
          description={null}
          defaultActionLabel={t("spreadsheet.cellShiftOk")}
          cancelActionLabel={t("spreadsheet.cellShiftCancel")}
          defaultActionTestId="spreadsheet-custom-views-add-ok"
          defaultActionDisabled={!customViewsDialog.adding.name.trim()}
          onDefaultAction={addCustomView}
          onCancel={() => setCustomViewsDialog((current) => (current ? { ...current, adding: null } : current))}
          returnFocusTarget={commandDialogReturnFocusTarget}
          className="sm:max-w-[380px]"
        >
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>{t("spreadsheet.customViews.name")}</span>
            <input
              type="text"
              autoFocus
              maxLength={255}
              value={customViewsDialog.adding.name}
              onChange={(event) =>
                setCustomViewsDialog((current) =>
                  current?.adding ? { ...current, adding: { ...current.adding, name: event.target.value } } : current,
                )
              }
              data-testid="spreadsheet-custom-views-add-name"
            />
          </label>
          <fieldset style={{ border: "none", padding: 0, margin: 0, display: "grid", gap: 6 }}>
            <legend className={dialogParts.label}>{t("spreadsheet.customViews.include")}</legend>
            {(["printSettings", "hiddenAndFilters"] as const).map((option) => (
              <label key={option} className={dialogParts.check}>
                <input
                  type="checkbox"
                  checked={customViewsDialog.adding![option]}
                  onChange={(event) =>
                    setCustomViewsDialog((current) =>
                      current?.adding
                        ? { ...current, adding: { ...current.adding, [option]: event.target.checked } }
                        : current,
                    )
                  }
                  data-testid={`spreadsheet-custom-views-add-${option}`}
                />
                {t(`spreadsheet.customViews.${option}`)}
              </label>
            ))}
          </fieldset>
        </OfficeDialogShell>
      ) : null
  );
}
