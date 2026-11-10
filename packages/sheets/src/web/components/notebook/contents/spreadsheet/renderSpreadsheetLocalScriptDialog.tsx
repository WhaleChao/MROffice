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
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetLocalScriptDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetLocalScriptDialogInputs {
  automationDialogReturnFocusRef: RefObject<HTMLElement | null>;
  closeLocalScriptEditor: () => void;
  runScript: () => Promise<void>;
  scriptDialogOpen: boolean;
  scriptEditorRef: RefObject<HTMLTextAreaElement | null>;
  scriptRunning: boolean;
  scriptStatus: { ok: boolean; message: string; } | null;
  scriptText: string;
  setScriptText: Dispatch<SetStateAction<string>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetLocalScriptDialog({
  automationDialogReturnFocusRef,
  closeLocalScriptEditor,
  runScript,
  scriptDialogOpen,
  scriptEditorRef,
  scriptRunning,
  scriptStatus,
  scriptText,
  setScriptText,
  t,
}: RenderSpreadsheetLocalScriptDialogInputs) {
  return (
       scriptDialogOpen ? (
        <OfficeDialogShell
          open
          title={t("spreadsheet.automate.localScriptTitle")}
          description={t("spreadsheet.automate.localScriptDescription")}
          defaultActionLabel={t("spreadsheet.scriptRun")}
          busyActionLabel={t("spreadsheet.scriptRunning")}
          cancelActionLabel={t("spreadsheet.macroCancel")}
          onDefaultAction={() => void runScript()}
          onCancel={closeLocalScriptEditor}
          defaultActionDisabled={scriptText.trim().length === 0}
          busy={scriptRunning}
          initialFocusRef={scriptEditorRef}
          returnFocusRef={automationDialogReturnFocusRef}
          className="sm:max-w-[640px]"
          testIdPrefix="spreadsheet-automation-script"
        >
          <textarea
            ref={scriptEditorRef}
            value={scriptText}
            onChange={(event) => setScriptText(event.target.value)}
            spellCheck={false}
            placeholder={`sheet.forEachCell(function (value, row, col) {\n  if (typeof value === "number") sheet.setCell(row, col, value * 2);\n});`}
            data-testid="spreadsheet-script-input"
            rows={12}
            className={`${dialogParts.mono} min-h-[240px] w-full resize-y`}
          />
          <div className={dialogParts.message} data-tone="warning">
            {t("spreadsheet.automate.localScriptBoundary")}
          </div>
          {scriptStatus ? (
            <div
              role="status"
              data-testid="spreadsheet-script-status"
              className={dialogParts.message}
              data-tone={scriptStatus.ok ? "success" : "error"}
            >
              {scriptStatus.message}
            </div>
          ) : null}
        </OfficeDialogShell>
      ) : null
  );
}
