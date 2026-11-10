"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { protectedRangeDisplayRef, normalizeProtectedRangeRef, type ProtectedEditRange } from "@mrnote/core/web/lib/spreadsheet/protected-ranges.ts";
import { spreadsheetNameFold } from "@mrnote/core/web/lib/spreadsheet/name-fold.ts";
import { type SheetModel } from "./spreadsheetSheetModel";
import { columnLabel } from "./spreadsheetRangeA1";
import { EDIT_RANGE_PASSWORD_KEPT } from "./spreadsheetWindowConstants";
import { type GridRect } from "@mrnote/core/web/lib/spreadsheet/grid-nav.ts";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetEditRangesDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetEditRangesDialogInputs {
  commandDialogReturnFocusTarget: () => HTMLElement | null;
  commitEditRanges: (ranges: ProtectedEditRange[]) => void;
  editRangesDialog: { ranges: ProtectedEditRange[]; selected: number | null; editor: { index: number | null; title: string; ref: string; password: string; confirm?: string; error?: string; } | null; } | null;
  selRect: GridRect;
  setEditRangesDialog: Dispatch<SetStateAction<{ ranges: ProtectedEditRange[]; selected: number | null; editor: { index: number | null; title: string; ref: string; password: string; confirm?: string; error?: string; } | null; } | null>>;
  sheet: SheetModel;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetEditRangesDialog({
  commandDialogReturnFocusTarget,
  commitEditRanges,
  editRangesDialog,
  selRect,
  setEditRangesDialog,
  sheet,
  t,
}: RenderSpreadsheetEditRangesDialogInputs) {
  return (
       editRangesDialog && !editRangesDialog.editor ? (
        // XLS-A028: Excel's Allow Users to Edit Ranges.
        <OfficeDialogShell
          open
          testIdPrefix="spreadsheet-edit-ranges"
          title={t("spreadsheet.editRanges.title")}
          description={t("spreadsheet.editRanges.listLabel")}
          defaultActionLabel={t("spreadsheet.cellShiftOk")}
          cancelActionLabel={t("spreadsheet.cellShiftCancel")}
          defaultActionTestId="spreadsheet-edit-ranges-ok"
          onDefaultAction={() => commitEditRanges(editRangesDialog.ranges)}
          onCancel={() => setEditRangesDialog(null)}
          footerStart={
            <>
              <button
                type="button"
                className={dialogParts.button}
                data-testid="spreadsheet-edit-ranges-new"
                onClick={() => {
                  let n = editRangesDialog.ranges.length + 1;
                  const taken = new Set(editRangesDialog.ranges.map((range) => spreadsheetNameFold(range.title)));
                  while (taken.has(`range${n}`)) n += 1;
                  setEditRangesDialog({
                    ...editRangesDialog,
                    editor: {
                      index: null,
                      title: `Range${n}`,
                      ref: protectedRangeDisplayRef(
                        normalizeProtectedRangeRef(
                          `${columnLabel(selRect.c1)}${selRect.r1 + 1}:${columnLabel(selRect.c2)}${selRect.r2 + 1}`,
                          sheet.name,
                        ) ?? "A1",
                      ),
                      password: "",
                    },
                  });
                }}
              >
                {t("spreadsheet.editRanges.new")}
              </button>
              <button
                type="button"
                className={dialogParts.button}
                disabled={editRangesDialog.selected === null}
                data-testid="spreadsheet-edit-ranges-modify"
                onClick={() => {
                  const index = editRangesDialog.selected;
                  const range = index === null ? undefined : editRangesDialog.ranges[index];
                  if (index === null || !range) return;
                  setEditRangesDialog({
                    ...editRangesDialog,
                    editor: {
                      index,
                      title: range.title,
                      ref: protectedRangeDisplayRef(range.ref),
                      password: range.passwordHash ? EDIT_RANGE_PASSWORD_KEPT : "",
                    },
                  });
                }}
              >
                {t("spreadsheet.editRanges.modify")}
              </button>
              <button
                type="button"
                className={dialogParts.button}
                disabled={editRangesDialog.selected === null}
                data-testid="spreadsheet-edit-ranges-delete"
                onClick={() => {
                  const ranges = editRangesDialog.ranges.filter((_, index) => index !== editRangesDialog.selected);
                  setEditRangesDialog({ ...editRangesDialog, ranges, selected: ranges.length ? 0 : null });
                }}
              >
                {t("spreadsheet.editRanges.delete")}
              </button>
            </>
          }
          returnFocusTarget={commandDialogReturnFocusTarget}
          className="sm:max-w-[460px]"
        >
          <table className="w-full text-left text-sm" data-testid="spreadsheet-edit-ranges-list">
            <thead>
              <tr>
                <th className="px-2 py-1 font-medium">{t("spreadsheet.editRanges.titleColumn")}</th>
                <th className="px-2 py-1 font-medium">{t("spreadsheet.editRanges.refColumn")}</th>
              </tr>
            </thead>
            <tbody>
              {editRangesDialog.ranges.map((range, index) => (
                <tr
                  key={range.title}
                  aria-selected={editRangesDialog.selected === index}
                  className={editRangesDialog.selected === index ? "bg-sky-100" : undefined}
                  data-testid={`spreadsheet-edit-ranges-row-${index}`}
                  onClick={() => setEditRangesDialog({ ...editRangesDialog, selected: index })}
                >
                  <td className="px-2 py-1">{range.title}</td>
                  <td className="px-2 py-1">{protectedRangeDisplayRef(range.ref)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </OfficeDialogShell>
      ) : null
  );
}
