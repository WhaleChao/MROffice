"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { type ProtectedEditRange } from "@mrnote/core/web/lib/spreadsheet/protected-ranges.ts";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetEditRangeConfirmDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetEditRangeConfirmDialogInputs {
  commandDialogReturnFocusTarget: () => HTMLElement | null;
  editRangesDialog: { ranges: ProtectedEditRange[]; selected: number | null; editor: { index: number | null; title: string; ref: string; password: string; confirm?: string; error?: string; } | null; } | null;
  saveEditRangeEditor: () => Promise<void>;
  setEditRangesDialog: Dispatch<SetStateAction<{ ranges: ProtectedEditRange[]; selected: number | null; editor: { index: number | null; title: string; ref: string; password: string; confirm?: string; error?: string; } | null; } | null>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetEditRangeConfirmDialog({
  commandDialogReturnFocusTarget,
  editRangesDialog,
  saveEditRangeEditor,
  setEditRangesDialog,
  t,
}: RenderSpreadsheetEditRangeConfirmDialogInputs) {
  return (
       editRangesDialog?.editor ? (
        // XLS-A028: New Range / Modify Range, then Confirm Password.
        <OfficeDialogShell
          open
          testIdPrefix="spreadsheet-edit-range-editor"
          title={
            editRangesDialog.editor.confirm !== undefined
              ? t("spreadsheet.editRanges.confirmTitle")
              : editRangesDialog.editor.index === null
                ? t("spreadsheet.editRanges.newTitle")
                : t("spreadsheet.editRanges.modifyTitle")
          }
          description={null}
          defaultActionLabel={t("spreadsheet.cellShiftOk")}
          cancelActionLabel={t("spreadsheet.cellShiftCancel")}
          defaultActionTestId="spreadsheet-edit-range-editor-ok"
          onDefaultAction={() => void saveEditRangeEditor()}
          onCancel={() => setEditRangesDialog({ ...editRangesDialog, editor: null })}
          returnFocusTarget={commandDialogReturnFocusTarget}
          className="sm:max-w-[380px]"
        >
          {editRangesDialog.editor.confirm !== undefined ? (
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.editRanges.reenter")}</span>
              <input
                type="password"
                autoFocus
                value={editRangesDialog.editor.confirm}
                onChange={(event) =>
                  setEditRangesDialog({
                    ...editRangesDialog,
                    editor: { ...editRangesDialog.editor!, confirm: event.target.value },
                  })
                }
                data-testid="spreadsheet-edit-range-confirm"
              />
            </label>
          ) : (
            <>
              {(["title", "ref", "password"] as const).map((field) => (
                <label key={field} className={dialogParts.field}>
                  <span className={dialogParts.label}>{t(`spreadsheet.editRanges.field.${field}`)}</span>
                  <input
                    type={field === "password" ? "password" : "text"}
                    autoFocus={field === "title"}
                    value={editRangesDialog.editor![field]}
                    onChange={(event) =>
                      setEditRangesDialog({
                        ...editRangesDialog,
                        editor: { ...editRangesDialog.editor!, [field]: event.target.value, error: undefined },
                      })
                    }
                    data-testid={`spreadsheet-edit-range-${field}`}
                  />
                </label>
              ))}
            </>
          )}
          {editRangesDialog.editor.error ? (
            <p role="alert" className={dialogParts.error} data-testid="spreadsheet-edit-range-error">
              {editRangesDialog.editor.error}
            </p>
          ) : null}
        </OfficeDialogShell>
      ) : null
  );
}
