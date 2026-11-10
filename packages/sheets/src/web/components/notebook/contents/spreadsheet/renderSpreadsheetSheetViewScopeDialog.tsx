"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { type SheetViewScopedAction } from "./spreadsheetSheetViews";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetSheetViewScopeDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetSheetViewScopeDialogInputs {
  resolveSheetViewScopePrompt: (choice: "me" | "everyone" | "cancel") => void;
  setSheetViewScopePrompt: Dispatch<SetStateAction<{ action: SheetViewScopedAction; choice: "me" | "everyone"; remember: boolean; } | null>>;
  sheetViewScopePrompt: { action: SheetViewScopedAction; choice: "me" | "everyone"; remember: boolean; } | null;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetSheetViewScopeDialog({
  resolveSheetViewScopePrompt,
  setSheetViewScopePrompt,
  sheetViewScopePrompt,
  t,
}: RenderSpreadsheetSheetViewScopeDialogInputs) {
  return (
      <OfficeDialogShell
        open={sheetViewScopePrompt !== null}
        title={t("spreadsheet.sheetViewScope.title")}
        description={t("spreadsheet.sheetViewScope.description")}
        defaultActionLabel={t("spreadsheet.sheetViewScope.apply")}
        cancelActionLabel={t("common.cancel")}
        onDefaultAction={() =>
          resolveSheetViewScopePrompt(sheetViewScopePrompt?.choice ?? "me")
        }
        onCancel={() => resolveSheetViewScopePrompt("cancel")}
        testIdPrefix="spreadsheet-sheet-view-scope"
      >
        <div
          role="radiogroup"
          aria-label={t("spreadsheet.sheetViewScope.title")}
          className={dialogParts.choices}
        >
          {(["me", "everyone"] as const).map((option) => {
            const checked = sheetViewScopePrompt?.choice === option;
            return (
              <button
                key={option}
                type="button"
                role="radio"
                aria-checked={checked}
                onClick={() =>
                  setSheetViewScopePrompt((current) =>
                    current ? { ...current, choice: option } : current,
                  )
                }
                data-testid={`spreadsheet-sheet-view-scope-${option}`}
                className={`${dialogParts.choice} grid gap-0.5 text-left`}
              >
                <span className="font-medium">
                  {option === "me"
                    ? t("spreadsheet.sheetViewScope.me")
                    : t("spreadsheet.sheetViewScope.everyone")}
                </span>
                <span className={dialogParts.hint}>
                  {option === "me"
                    ? t("spreadsheet.sheetViewScope.meHint")
                    : t("spreadsheet.sheetViewScope.everyoneHint")}
                </span>
              </button>
            );
          })}
        </div>
        <label className={dialogParts.check}>
          <input
            type="checkbox"
            checked={sheetViewScopePrompt?.remember ?? false}
            onChange={(event) =>
              setSheetViewScopePrompt((current) =>
                current ? { ...current, remember: event.target.checked } : current,
              )
            }
            data-testid="spreadsheet-sheet-view-scope-remember"
          />
          {t("spreadsheet.sheetViewScope.remember")}
        </label>
      </OfficeDialogShell>
  );
}
