"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { ERROR_CHECK_RULES } from "../../../../lib/spreadsheet/error-checking";
import {
  readErrorCheckingOptions,
  writeErrorCheckingOptions,
} from "../../../../lib/spreadsheet/error-checking-options";
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

interface SpreadsheetErrorCheckingOptionsDialogProps {
  /** Whether any cell of the workbook carries an Ignore Error choice. */
  readonly hasIgnoredErrors: boolean;
  /** Why Reset Ignored Errors is refused here (a read-only page). */
  readonly resetDisabledReason?: string;
  readonly onResetIgnoredErrors: () => void;
  readonly onClose: () => void;
}

/**
 * Error Checking Options (the last command of the error button's menu, which
 * in Excel opens File > Options > Formulas). Microsoft Support, "Detect
 * errors in formulas", "Turn error checking rules on or off": "Enable
 * background error checking", the "Excel checking rules" tick boxes, and
 * Reset Ignored Errors.
 *
 * The switches are a per-user preference applied on OK
 * (`lib/spreadsheet/error-checking-options`). Reset Ignored Errors acts at
 * once, as Excel's button does, and is a workbook write the window owns (one
 * commit, one Undo): it brings back every triangle silenced with Ignore Error,
 * on every sheet. "Indicate errors using this color" is left out: the
 * triangle's green is the theme's, not a stored choice.
 *
 * Mounted only while open, so each opening starts from the stored choice.
 */
export function SpreadsheetErrorCheckingOptionsDialog({
  hasIgnoredErrors,
  resetDisabledReason,
  onResetIgnoredErrors,
  onClose,
}: SpreadsheetErrorCheckingOptionsDialogProps) {
  const t = useTranslations("console-notebooks");
  const [options, setOptions] = useState(readErrorCheckingOptions);

  return (
    <OfficeDialogShell
      open
      title={t("spreadsheet.errorCheckingOptions.title")}
      description={t("spreadsheet.errorCheckingOptions.description")}
      defaultActionLabel={t("spreadsheet.fmtOk")}
      cancelActionLabel={t("common.cancel")}
      onDefaultAction={() => {
        writeErrorCheckingOptions(options);
        onClose();
      }}
      onCancel={onClose}
      className="sm:max-w-[460px]"
      testIdPrefix="spreadsheet-error-checking-options"
    >
      <fieldset className="grid gap-2">
        <legend>
          {t("spreadsheet.errorCheckingOptions.checkingHeading")}
        </legend>
        <label className={officeDialogParts.check}>
          <input
            type="checkbox"
            checked={options.background}
            onChange={(event) =>
              setOptions({ ...options, background: event.target.checked })
            }
            data-testid="spreadsheet-error-checking-options-background"
          />
          {t("spreadsheet.errorCheckingOptions.background")}
        </label>
        <div>
          <button
            type="button"
            className={officeDialogParts.compactButton}
            disabled={!hasIgnoredErrors || Boolean(resetDisabledReason)}
            title={resetDisabledReason}
            onClick={onResetIgnoredErrors}
            data-testid="spreadsheet-error-checking-options-reset-ignored"
          >
            <SpreadsheetDialogIcon name="reset" />
            {t("spreadsheet.errorCheckingOptions.resetIgnored")}
          </button>
        </div>
      </fieldset>
      <fieldset className="grid gap-1.5">
        <legend>
          {t("spreadsheet.errorCheckingOptions.rulesHeading")}
        </legend>
        <div className={`${officeDialogParts.well} grid gap-1.5 px-3 py-2`}>
          {ERROR_CHECK_RULES.map((rule) => (
            <label
              key={rule}
              className={officeDialogParts.check}
            >
              <input
                type="checkbox"
                checked={options.rules[rule]}
                onChange={(event) =>
                  setOptions({
                    ...options,
                    rules: { ...options.rules, [rule]: event.target.checked },
                  })
                }
                data-testid={`spreadsheet-error-checking-options-rule-${rule}`}
              />
              {t(`spreadsheet.errorCheckingOptions.rule.${rule}`)}
            </label>
          ))}
        </div>
      </fieldset>
    </OfficeDialogShell>
  );
}
