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
import { ProtectionPasswordConfirmFields } from "./SpreadsheetProtectionPasswordDialog";
import { type WorkbookStructureProtection, workbookProtectionActive } from "@mrnote/core/web/lib/spreadsheet/protection.ts";
import { type WorkbookProtectionDialogAuthority, type WorkbookProtectionMutationReceipt } from "./spreadsheetMutationReceipts";
import { type Dispatch, type SetStateAction } from "react";
import { type ResolvedSpreadsheetVisibleControl } from "../../../../lib/spreadsheet/visible-command-authority";
import { type SheetProtectionPassword } from "@mrnote/core/web/lib/spreadsheet/protection.ts";

/** What renderSpreadsheetWorkbookProtectionDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetWorkbookProtectionDialogInputs {
  confirmWorkbookProtection: () => Promise<boolean>;
  setWorkbookProtectionConfirmDraft: Dispatch<SetStateAction<string | null>>;
  setWorkbookProtectionDialogAuthority: Dispatch<SetStateAction<WorkbookProtectionDialogAuthority | null>>;
  setWorkbookProtectionError: Dispatch<SetStateAction<string | null>>;
  setWorkbookProtectionPasswordDraft: Dispatch<SetStateAction<string>>;
  setWorkbookProtectionStructure: Dispatch<SetStateAction<boolean>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  visibleCommandState: (controlId: string) => ResolvedSpreadsheetVisibleControl;
  workbookProtectionBusy: boolean;
  workbookProtectionConfirmDraft: string | null;
  workbookProtectionDialogAuthority: WorkbookProtectionDialogAuthority | null;
  workbookProtectionError: string | null;
  workbookProtectionMutationCommandRef: RefObject<WorkbookProtectionMutationReceipt | null>;
  workbookProtectionPasswordDraft: string;
  workbookProtectionPasswordInputRef: RefObject<HTMLInputElement | null>;
  workbookProtectionStructure: boolean;
  workbookProtectionUnsupportedReason: string;
  workbookProtectionVerifier: SheetProtectionPassword | undefined;
  workbookStructureProtection: WorkbookStructureProtection | undefined;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetWorkbookProtectionDialog({
  confirmWorkbookProtection,
  setWorkbookProtectionConfirmDraft,
  setWorkbookProtectionDialogAuthority,
  setWorkbookProtectionError,
  setWorkbookProtectionPasswordDraft,
  setWorkbookProtectionStructure,
  t,
  visibleCommandState,
  workbookProtectionBusy,
  workbookProtectionConfirmDraft,
  workbookProtectionDialogAuthority,
  workbookProtectionError,
  workbookProtectionMutationCommandRef,
  workbookProtectionPasswordDraft,
  workbookProtectionPasswordInputRef,
  workbookProtectionStructure,
  workbookProtectionUnsupportedReason,
  workbookProtectionVerifier,
  workbookStructureProtection,
}: RenderSpreadsheetWorkbookProtectionDialogInputs) {
  // Structure, windows or both: the dialog is Unprotect Workbook.
  const workbookProtectionOn = workbookProtectionActive(workbookStructureProtection);
  return (
      <OfficeDialogShell
        open={workbookProtectionDialogAuthority !== null}
        title={
          workbookProtectionOn
            ? t("spreadsheet.uiCopy.unprotectWorkbook")
            : workbookProtectionConfirmDraft !== null
              ? t("spreadsheet.protectionPassword.confirmTitle")
              : t("spreadsheet.uiCopy.protectWorkbookStructure")
        }
        description={
          workbookProtectionOn
            ? t("spreadsheet.uiCopy.unprotectingAllowsSheetsToBeAddedDeletedRenamedMoved")
            : workbookProtectionConfirmDraft !== null
              ? t("spreadsheet.protectionPassword.reenter")
              : t("spreadsheet.protectWorkbookDialog.description")
        }
        defaultActionLabel={t("spreadsheet.fmtOk")}
        cancelActionLabel={t("common.cancel")}
        defaultActionDisabled={
          !visibleCommandState(
            "XLSCTL.REVIEW.PROTECT.WORKBOOK.MRAI.APPLY",
          ).enabled
        }
        busy={workbookProtectionBusy}
        busyActionLabel={
          workbookProtectionOn
            ? t("spreadsheet.protectionPassword.checking")
            : t("spreadsheet.protectionPassword.applying")
        }
        onDefaultAction={confirmWorkbookProtection}
        onCancel={() => {
          // Cancel on Confirm Password goes back to Protect Workbook.
          if (workbookProtectionConfirmDraft !== null) {
            setWorkbookProtectionConfirmDraft(null);
            setWorkbookProtectionError(null);
            return;
          }
          workbookProtectionMutationCommandRef.current = null;
          setWorkbookProtectionDialogAuthority(null);
          setWorkbookProtectionPasswordDraft("");
          setWorkbookProtectionError(null);
        }}
        initialFocusRef={
          workbookStructureProtection?.nativeBlocker
            ? undefined
            : workbookProtectionPasswordInputRef
        }
        initialFocus={
          workbookStructureProtection?.nativeBlocker ? "default" : undefined
        }
        testIdPrefix="spreadsheet-workbook-protection"
      >
        {workbookStructureProtection?.nativeBlocker ? (
          <p
            role="status"
            className={dialogParts.message}
            data-tone="warning"
            data-testid="spreadsheet-workbook-protection-native-blocked"
          >
            {workbookProtectionUnsupportedReason}
          </p>
        ) : !workbookProtectionOn &&
          workbookProtectionConfirmDraft !== null ? (
          <ProtectionPasswordConfirmFields
            value={workbookProtectionConfirmDraft}
            error={workbookProtectionError}
            onChange={(value) => {
              setWorkbookProtectionConfirmDraft(value);
              setWorkbookProtectionError(null);
            }}
            testIdPrefix="spreadsheet-workbook-protection"
          />
        ) : workbookProtectionOn && !workbookProtectionVerifier ? (
          <p>
            {t("spreadsheet.uiCopy.thisWorkbookUsesPasswordlessStructureProtection")}
          </p>
        ) : (
          <>
            {workbookProtectionOn ? null : (
              <fieldset className="grid gap-1.5">
                <legend>{t("spreadsheet.protectWorkbookDialog.protectFor")}</legend>
                <label className={dialogParts.check}>
                  <input
                    type="checkbox"
                    checked={workbookProtectionStructure}
                    onChange={(event) =>
                      setWorkbookProtectionStructure(event.target.checked)
                    }
                    data-testid="spreadsheet-workbook-protection-structure"
                  />
                  {t("spreadsheet.protectWorkbookDialog.structure")}
                </label>
              </fieldset>
            )}
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>
                {workbookProtectionOn
                  ? t("spreadsheet.protectionPassword.passwordLabel")
                  : t("spreadsheet.protectWorkbookDialog.password")}
              </span>
              <input
                ref={workbookProtectionPasswordInputRef}
                type="password"
                autoComplete={
                  workbookProtectionOn
                    ? "current-password"
                    : "new-password"
                }
                maxLength={255}
                value={workbookProtectionPasswordDraft}
                onChange={(event) => {
                  setWorkbookProtectionPasswordDraft(event.target.value);
                  setWorkbookProtectionError(null);
                }}
                aria-invalid={Boolean(workbookProtectionError)}
                aria-describedby={
                  workbookProtectionError
                    ? "spreadsheet-workbook-protection-error"
                    : undefined
                }
                data-testid="spreadsheet-workbook-protection-password"
              />
            </label>
            {workbookProtectionOn ? null : (
              <p className={dialogParts.hint}>
                {t("spreadsheet.uiCopy.workbookStructureProtectionDoesNotEncryptTheFileOr")}
              </p>
            )}
          </>
        )}
        {workbookProtectionError && workbookProtectionConfirmDraft === null ? (
          <p
            id="spreadsheet-workbook-protection-error"
            role="alert"
            className={dialogParts.error}
            data-testid="spreadsheet-workbook-protection-error"
          >
            {workbookProtectionError}
          </p>
        ) : null}
      </OfficeDialogShell>
  );
}
