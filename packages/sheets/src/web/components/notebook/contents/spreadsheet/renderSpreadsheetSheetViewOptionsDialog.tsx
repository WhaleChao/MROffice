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
import { normalizeNamedSheetViewName, MAX_NAMED_SHEET_VIEWS } from "../../../../lib/spreadsheet/sheet-views";
import { type SheetModel } from "./spreadsheetSheetModel";
import { type SheetViewDialogState } from "./spreadsheetSheetViews";
import { type SpreadsheetNamedSheetView } from "../../../../lib/spreadsheet/sheet-views";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetSheetViewOptionsDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetSheetViewOptionsDialogInputs {
  commitSheetViewDialog: () => void;
  deleteSelectedSheetView: () => void;
  duplicateSelectedSheetView: () => void;
  selectedSheetViewOption: SpreadsheetNamedSheetView | null;
  setSheetViewDialog: Dispatch<SetStateAction<SheetViewDialogState | null>>;
  sheet: SheetModel;
  sheetViewDialog: SheetViewDialogState | null;
  sheetViewDialogInputRef: RefObject<HTMLInputElement | null>;
  sheetViewReturnFocusRef: RefObject<HTMLElement | null>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  webDataReviewViewLabels: { getTransform: string; getData: string; fromTextCsvValues: string; fromTableRange: string; dataFromPicture: string; dataFromPowerBi: string; queriesConnections: string; connectionProperties: string; refresh: string; refreshAll: string; workbookLinks: string; dataTypes: string; stocks: string; currencies: string; sortFilter: string; customSort: string; clear: string; reapply: string; noActiveFilter: string; dataTools: string; analyzeData: string; analyzeDataUnavailable: string; outline: string; comments: string; previous: string; next: string; changes: string; proofing: string; protection: string; notes: string; sheetViews: string; defaultView: string; temporaryView: string; exitView: string; saveView: string; newView: string; viewOptions: string; namedViewsUnavailable: string; sheetViewRequired: string; temporarySheetViewRequired: string; sheetViewUnsupported: string; zoom: string; window: string; fullScreen: string; fullScreenUnavailable: string; show: string; headings: string; gridlines: string; formulaBar: string; navigation: string; dataTypeIcons: string; focusCell: string; focusCellColor: string; view: string; immersiveReader: string; immersiveReaderUnavailable: string; };
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetSheetViewOptionsDialog({
  commitSheetViewDialog,
  deleteSelectedSheetView,
  duplicateSelectedSheetView,
  selectedSheetViewOption,
  setSheetViewDialog,
  sheet,
  sheetViewDialog,
  sheetViewDialogInputRef,
  sheetViewReturnFocusRef,
  t,
  webDataReviewViewLabels,
}: RenderSpreadsheetSheetViewOptionsDialogInputs) {
  return (
      <OfficeDialogShell
        open={sheetViewDialog !== null}
        title={
          sheetViewDialog?.mode === "options"
            ? t("spreadsheet.view.sheetViewOptions")
            : t("spreadsheet.uiCopy.keepSheetView")
        }
        description={
          sheetViewDialog?.mode === "options"
            ? t("spreadsheet.uiCopy.renameDuplicateOrDeleteASavedFilterAndSort")
            : t("spreadsheet.uiCopy.nameTheCurrentTemporaryFilterView")
        }
        defaultActionLabel={
          sheetViewDialog?.mode === "options"
            ? t("code.rename.submit")
            : t("spreadsheet.uiCopy.keep")
        }
        cancelActionLabel={t("common.cancel")}
        defaultActionDisabled={
          !sheetViewDialog ||
          Boolean(selectedSheetViewOption?.nativeBlocker) ||
          !normalizeNamedSheetViewName(sheetViewDialog.name)
        }
        onDefaultAction={commitSheetViewDialog}
        onCancel={() => setSheetViewDialog(null)}
        initialFocusRef={sheetViewDialogInputRef}
        returnFocusRef={sheetViewReturnFocusRef}
        testIdPrefix="spreadsheet-sheet-view"
      >
        {sheetViewDialog?.mode === "options" ? (
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>{t("toolbar.ribbonView")}</span>
            <select
              value={sheetViewDialog.selectedId}
              onChange={(event) => {
                const selected = (sheet.namedSheetViews ?? []).find(
                  (view) => view.id === event.target.value,
                );
                if (!selected) return;
                setSheetViewDialog({
                  mode: "options",
                  selectedId: selected.id,
                  name: selected.name,
                });
              }}
              data-testid="spreadsheet-sheet-view-select"
            >
              {(sheet.namedSheetViews ?? []).map((view) => (
                <option key={view.id} value={view.id}>
                  {view.name}
                </option>
              ))}
            </select>
          </label>
        ) : null}
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>{t("presentation.master.name")}</span>
          <input
            ref={sheetViewDialogInputRef}
            value={sheetViewDialog?.name ?? ""}
            maxLength={254}
            disabled={Boolean(selectedSheetViewOption?.nativeBlocker)}
            onChange={(event) =>
              setSheetViewDialog((current) =>
                current
                  ? { ...current, name: event.target.value, error: undefined }
                  : current,
              )
            }
            data-testid="spreadsheet-sheet-view-name"
          />
        </label>
        {sheetViewDialog?.error ? (
          <p
            role="alert"
            data-testid="spreadsheet-sheet-view-error"
            className={dialogParts.error}
          >
            {sheetViewDialog.error === "duplicate"
              ? t("spreadsheet.uiCopy.viewNamesMustBeUniqueWithinTheWorksheet")
              : t("spreadsheet.uiCopy.enterAValidViewName")}
          </p>
        ) : null}
        {selectedSheetViewOption?.nativeBlocker ? (
          <p
            role="alert"
            data-testid="spreadsheet-sheet-view-native-blocker"
            className={dialogParts.message}
            data-tone="warning"
          >
            {webDataReviewViewLabels.sheetViewUnsupported}
          </p>
        ) : null}
        {sheetViewDialog?.mode === "options" ? (
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              disabled={
                Boolean(selectedSheetViewOption?.nativeBlocker) ||
                (sheet.namedSheetViews?.length ?? 0) >= MAX_NAMED_SHEET_VIEWS
              }
              onClick={duplicateSelectedSheetView}
              data-testid="spreadsheet-sheet-view-duplicate"
              className={dialogParts.compactButton}
            >
              {t("spreadsheet.uiCopy.duplicate")}
            </button>
            <button
              type="button"
              // A preserve-only view can still be deleted: the write-back
              // drops its element and keeps the other views' as they are.
              disabled={Boolean(sheet.namedSheetViewsNativeBlocker)}
              onClick={deleteSelectedSheetView}
              data-testid="spreadsheet-sheet-view-delete"
              className={dialogParts.compactButton}
            >
              {t("presentation.contextMenu.delete")}
            </button>
          </div>
        ) : null}
      </OfficeDialogShell>
  );
}
