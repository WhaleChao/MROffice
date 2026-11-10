"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { type NamedRange } from "@mrnote/core/web/lib/spreadsheet/named-ranges.ts";
import { type SheetModel } from "./spreadsheetSheetModel";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetPasteNamesDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetPasteNamesDialogInputs {
  commandDialogReturnFocusTarget: () => HTMLElement | null;
  insertNameIntoFormula: (name: string) => void;
  pasteNameList: () => void;
  pasteNamesDraft: { selected: string | null; } | null;
  setPasteNamesDraft: Dispatch<SetStateAction<{ selected: string | null; } | null>>;
  sheet: SheetModel;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  workbookNamedRanges: NamedRange[];
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetPasteNamesDialog({
  commandDialogReturnFocusTarget,
  insertNameIntoFormula,
  pasteNameList,
  pasteNamesDraft,
  setPasteNamesDraft,
  sheet,
  t,
  workbookNamedRanges,
}: RenderSpreadsheetPasteNamesDialogInputs) {
  return (
       pasteNamesDraft ? (
        // XLS-A023: Use in Formula ▸ Paste Names….
        <OfficeDialogShell
          open
          testIdPrefix="spreadsheet-paste-names"
          title={t("spreadsheet.pasteNamesTitle")}
          description={null}
          defaultActionLabel={t("spreadsheet.cellShiftOk")}
          cancelActionLabel={t("spreadsheet.cellShiftCancel")}
          defaultActionTestId="spreadsheet-paste-names-ok"
          defaultActionDisabled={!pasteNamesDraft.selected}
          onDefaultAction={() => {
            const selected = pasteNamesDraft.selected;
            setPasteNamesDraft(null);
            if (selected) insertNameIntoFormula(selected);
          }}
          onCancel={() => setPasteNamesDraft(null)}
          footerStart={
            <button
              type="button"
              className={dialogParts.button}
              data-testid="spreadsheet-paste-names-list"
              onClick={pasteNameList}
            >
              {t("spreadsheet.pasteNamesList")}
            </button>
          }
          returnFocusTarget={commandDialogReturnFocusTarget}
          className="sm:max-w-[320px]"
        >
          <div role="listbox" aria-label={t("spreadsheet.pasteNamesTitle")} className={`${dialogParts.list} max-h-60`}>
            {workbookNamedRanges
              .filter(
                (entry) =>
                  !entry.hidden &&
                  (entry.scope !== "worksheet" || entry.scopeSheetId === sheet.id),
              )
              .map((entry) => (
                <div
                  key={`${entry.scope ?? "workbook"}:${entry.scopeSheetId ?? ""}:${entry.name}`}
                  role="option"
                  tabIndex={0}
                  aria-selected={pasteNamesDraft.selected === entry.name}
                  className={dialogParts.listItem}
                  data-testid={`spreadsheet-paste-names-${entry.name}`}
                  onClick={() => setPasteNamesDraft({ selected: entry.name })}
                  onDoubleClick={() => {
                    setPasteNamesDraft(null);
                    insertNameIntoFormula(entry.name);
                  }}
                  onKeyDown={(event) => {
                    if (event.key === " ") {
                      event.preventDefault();
                      setPasteNamesDraft({ selected: entry.name });
                    }
                  }}
                >
                  {entry.name}
                </div>
              ))}
          </div>
        </OfficeDialogShell>
      ) : null
  );
}
