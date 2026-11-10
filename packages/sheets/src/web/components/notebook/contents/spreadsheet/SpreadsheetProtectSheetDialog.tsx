"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  EXCEL_DEFAULT_SHEET_PROTECTION_PERMISSIONS,
  SHEET_PROTECTION_PERMISSION_ORDER,
  sheetProtectionOptionsFrom,
  type SheetProtectionOptions,
  type SheetProtectionPermission,
  type SheetProtectionPermissions,
} from "@mrnote/core/web/lib/spreadsheet/protection.ts";
import { createProtectionPasswordHash } from "@mrnote/core/web/lib/spreadsheet/protection-password.ts";

import { ProtectionPasswordConfirmFields } from "./SpreadsheetProtectionPasswordDialog";

interface SpreadsheetProtectSheetDialogProps {
  /** The persisted options (permissions + Unprotect verifier), or undefined
   *  for Excel's defaults with no password. */
  readonly onConfirm: (options: SheetProtectionOptions | undefined) => void;
  readonly onCancel: () => void;
}

const PERMISSION_LABEL_KEYS: Record<SheetProtectionPermission, string> = {
  selectLockedCells: "spreadsheet.protectSheetDialog.selectLockedCells",
  selectUnlockedCells: "spreadsheet.protectSheetDialog.selectUnlockedCells",
  formatCells: "spreadsheet.protectSheetDialog.formatCells",
  formatColumns: "spreadsheet.protectSheetDialog.formatColumns",
  formatRows: "spreadsheet.protectSheetDialog.formatRows",
  insertColumns: "spreadsheet.protectSheetDialog.insertColumns",
  insertRows: "spreadsheet.protectSheetDialog.insertRows",
  insertHyperlinks: "spreadsheet.protectSheetDialog.insertHyperlinks",
  deleteColumns: "spreadsheet.protectSheetDialog.deleteColumns",
  deleteRows: "spreadsheet.protectSheetDialog.deleteRows",
  sort: "spreadsheet.protectSheetDialog.sort",
  autoFilter: "spreadsheet.protectSheetDialog.autoFilter",
  pivotTables: "spreadsheet.protectSheetDialog.pivotTables",
  editObjects: "spreadsheet.protectSheetDialog.editObjects",
  editScenarios: "spreadsheet.protectSheetDialog.editScenarios",
};

const PERMISSION_TEST_IDS: Record<SheetProtectionPermission, string> = {
  selectLockedCells: "select-locked",
  selectUnlockedCells: "select-unlocked",
  formatCells: "format-cells",
  formatColumns: "format-columns",
  formatRows: "format-rows",
  insertColumns: "insert-columns",
  insertRows: "insert-rows",
  insertHyperlinks: "insert-hyperlinks",
  deleteColumns: "delete-columns",
  deleteRows: "delete-rows",
  sort: "sort",
  autoFilter: "autofilter",
  pivotTables: "pivot-tables",
  editObjects: "edit-objects",
  editScenarios: "edit-scenarios",
};

















export function SpreadsheetProtectSheetDialog({
  onConfirm,
  onCancel,
}: SpreadsheetProtectSheetDialogProps) {
  const t = useTranslations("console-notebooks");
  const [protectContents, setProtectContents] = useState(true);
  const [permissions, setPermissions] = useState<SheetProtectionPermissions>({
    ...EXCEL_DEFAULT_SHEET_PROTECTION_PERMISSIONS,
  });
  const [password, setPassword] = useState("");
  const [confirmation, setConfirmation] = useState("");
  const [step, setStep] = useState<"options" | "confirm">("options");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const passwordRef = useRef<HTMLInputElement>(null);
  const confirmRef = useRef<HTMLInputElement>(null);

  const toggle = (permission: SheetProtectionPermission, checked: boolean) => {
    setPermissions((current) => {
      const next = { ...current, [permission]: checked };
      if (permission === "selectUnlockedCells" && !checked) {
        next.selectLockedCells = false;
      }
      if (permission === "selectLockedCells" && checked) {
        next.selectUnlockedCells = true;
      }
      return next;
    });
  };

  const confirm = () => {
    if (busy || !protectContents) return;
    if (step === "options") {
      if (password.length === 0) {
        onConfirm(sheetProtectionOptionsFrom(permissions));
        return;
      }
      setConfirmation("");
      setError(null);
      setStep("confirm");
      return;
    }
    if (confirmation !== password) {
      setError(t("spreadsheet.protectionPassword.mismatch"));
      return;
    }
    setBusy(true);
    void createProtectionPasswordHash(password)
      .then((hash) => {
        setPassword("");
        setConfirmation("");
        onConfirm(sheetProtectionOptionsFrom(permissions, { kind: "hash", ...hash }));
      })
      .finally(() => setBusy(false));
  };

  const cancel = () => {
    if (step === "confirm") {
      setStep("options");
      setConfirmation("");
      setError(null);
      return;
    }
    onCancel();
  };

  return (
    <OfficeDialogShell
      open
      title={
        step === "confirm"
          ? t("spreadsheet.protectionPassword.confirmTitle")
          : t("spreadsheet.protectSheetDialog.title")
      }
      description={
        step === "confirm"
          ? t("spreadsheet.protectionPassword.reenter")
          : t("spreadsheet.protectSheetDialog.description")
      }
      defaultActionLabel={t("spreadsheet.protectSheetDialog.ok")}
      cancelActionLabel={t("common.cancel")}
      defaultActionDisabled={!protectContents}
      busy={busy}
      busyActionLabel={t("spreadsheet.protectionPassword.applying")}
      onDefaultAction={confirm}
      onCancel={cancel}
      initialFocusRef={step === "confirm" ? confirmRef : passwordRef}
      className="sm:max-w-[420px]"
      testIdPrefix="spreadsheet-protect-sheet"
    >
      {step === "confirm" ? (
        <ProtectionPasswordConfirmFields
          inputRef={confirmRef}
          value={confirmation}
          error={error}
          onChange={(value) => {
            setConfirmation(value);
            setError(null);
          }}
          testIdPrefix="spreadsheet-protect-sheet"
        />
      ) : (
        <>
          <label className={officeDialogParts.check}>
            <input
              type="checkbox"
              checked={protectContents}
              onChange={(event) => setProtectContents(event.target.checked)}
              data-testid="spreadsheet-protect-sheet-contents"
            />
            {t("spreadsheet.protectSheetDialog.protectContents")}
          </label>
          <label className={officeDialogParts.field}>
            <span className={officeDialogParts.label}>
              {t("spreadsheet.protectSheetDialog.password")}
            </span>
            <input
              ref={passwordRef}
              type="password"
              autoComplete="new-password"
              maxLength={255}
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              data-testid="spreadsheet-protect-sheet-password"
            />
          </label>
          <fieldset className="grid gap-1.5">
            <legend>{t("spreadsheet.protectSheetDialog.allowHeading")}</legend>
            <div
              className={`${officeDialogParts.well} grid max-h-[220px] gap-1.5 overflow-y-auto px-3 py-2`}
            >
              {SHEET_PROTECTION_PERMISSION_ORDER.map((permission) => (
                <label key={permission} className={officeDialogParts.check}>
                  <input
                    type="checkbox"
                    checked={permissions[permission]}
                    onChange={(event) => toggle(permission, event.target.checked)}
                    data-testid={`spreadsheet-protect-sheet-allow-${PERMISSION_TEST_IDS[permission]}`}
                  />
                  {t(PERMISSION_LABEL_KEYS[permission])}
                </label>
              ))}
            </div>
          </fieldset>
        </>
      )}
    </OfficeDialogShell>
  );
}
