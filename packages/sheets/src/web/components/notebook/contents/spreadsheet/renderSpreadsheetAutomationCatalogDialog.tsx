"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type RefObject } from "react";
import { Trash2 } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { SpreadsheetInsertIcon } from "./SpreadsheetInsertRibbonIcons";
import { type FreezeAction } from "@mrnote/core/web/lib/spreadsheet/freeze-panes.ts";
import { macroReferenceMode, type SpreadsheetMacro } from "../../../../lib/spreadsheet/macros";
import { type ClearVariant } from "../../../../lib/spreadsheet/block-clear";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetAutomationCatalogDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetAutomationCatalogDialogInputs {
  applyFreezeAction: (action: FreezeAction) => void;
  automationDialogReturnFocusRef: RefObject<HTMLElement | null>;
  clearSelectionBy: (variant: ClearVariant) => void;
  deleteMacro: (name: string) => void;
  freezeActive: boolean;
  homeNoHyperlinksReason: string;
  homeSelectionFacets: { comments: boolean; hyperlinks: boolean; };
  macros: SpreadsheetMacro[];
  manualHiddenCols: Set<number>;
  manualHiddenRows: Set<number>;
  openLocalScriptEditor: (template?: string, preserveReturnFocus?: boolean) => void;
  runAutomationCountBlankRows: () => void;
  runAutomationUnhideAll: () => void;
  runMacro: (name: string) => void;
  setAutomationCatalogOpen: Dispatch<SetStateAction<boolean>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetAutomationCatalogDialog({
  applyFreezeAction,
  automationDialogReturnFocusRef,
  clearSelectionBy,
  deleteMacro,
  freezeActive,
  homeNoHyperlinksReason,
  homeSelectionFacets,
  macros,
  manualHiddenCols,
  manualHiddenRows,
  openLocalScriptEditor,
  runAutomationCountBlankRows,
  runAutomationUnhideAll,
  runMacro,
  setAutomationCatalogOpen,
  t,
}: RenderSpreadsheetAutomationCatalogDialogInputs) {
  return (
        <OfficeDialogShell
          open
          title={t("spreadsheet.automate.catalogTitle")}
          description={t("spreadsheet.automate.catalogDescription")}
          defaultActionLabel={t("spreadsheet.automate.newLocalScript")}
          cancelActionLabel={t("common.close")}
          onDefaultAction={() => openLocalScriptEditor("", true)}
          onCancel={() => setAutomationCatalogOpen(false)}
          returnFocusRef={automationDialogReturnFocusRef}
          className="sm:max-w-[620px]"
          testIdPrefix="spreadsheet-automation-catalog"
        >
          <div
            className={`${dialogParts.choices} grid-cols-1 sm:grid-cols-2`}
            data-testid="spreadsheet-automation-catalog-actions"
          >
            <button
              type="button"
              disabled={
                manualHiddenRows.size === 0 && manualHiddenCols.size === 0
              }
              title={
                manualHiddenRows.size === 0 && manualHiddenCols.size === 0
                  ? t("spreadsheet.automate.nothingHidden")
                  : undefined
              }
              onClick={() => {
                runAutomationUnhideAll();
                setAutomationCatalogOpen(false);
              }}
              data-testid="spreadsheet-automation-catalog-unhide-all"
              className={`${dialogParts.choice} flex items-center gap-3 text-left font-medium`}
            >
              <SpreadsheetInsertIcon name="unhideAll" size={24} />
              {t("spreadsheet.automate.unhideAll")}
            </button>
            <button
              type="button"
              onClick={() => {
                applyFreezeAction(
                  freezeActive ? "unfreeze" : "freezePanes",
                );
                setAutomationCatalogOpen(false);
              }}
              data-testid="spreadsheet-automation-catalog-freeze"
              className={`${dialogParts.choice} flex items-center gap-3 text-left font-medium`}
            >
              <SpreadsheetInsertIcon name="freezeSelection" size={24} />
              {t(
                freezeActive
                  ? "spreadsheet.unfreeze"
                  : "spreadsheet.automate.freezeSelection",
              )}
            </button>
            <button
              type="button"
              disabled={!homeSelectionFacets.hyperlinks}
              title={
                homeSelectionFacets.hyperlinks
                  ? undefined
                  : homeNoHyperlinksReason
              }
              onClick={() => {
                clearSelectionBy("hyperlinks");
                setAutomationCatalogOpen(false);
              }}
              data-testid="spreadsheet-automation-catalog-clear-links"
              className={`${dialogParts.choice} flex items-center gap-3 text-left font-medium`}
            >
              <SpreadsheetInsertIcon name="clearLinks" size={24} />
              {t("spreadsheet.automate.clearHyperlinks")}
            </button>
            <button
              type="button"
              onClick={() => {
                runAutomationCountBlankRows();
                setAutomationCatalogOpen(false);
              }}
              data-testid="spreadsheet-automation-catalog-count-blank-rows"
              className={`${dialogParts.choice} flex items-center gap-3 text-left font-medium`}
            >
              <SpreadsheetInsertIcon name="countBlankRows" size={24} />
              {t("spreadsheet.automate.countBlankRows")}
            </button>
          </div>
          <section
            className={dialogParts.group}
            aria-label={t("spreadsheet.automate.recordedActionsSection")}
            data-testid="spreadsheet-automation-recorded-actions"
          >
            <h3 className={dialogParts.sectionTitle}>
              {t("spreadsheet.automate.recordedActionsSection")}
            </h3>
            {macros.length === 0 ? (
              <p className={dialogParts.hint}>
                {t("spreadsheet.macrosEmpty")}
              </p>
            ) : (
              <div className={dialogParts.list}>
                {macros.map((macro) => (
                  <div
                    key={macro.name}
                    className="flex min-w-0 items-center gap-2"
                    data-testid={`spreadsheet-automation-recorded-${macro.name}`}
                  >
                    <button
                      type="button"
                      onClick={() => {
                        runMacro(macro.name);
                        setAutomationCatalogOpen(false);
                      }}
                      title={
                        macroReferenceMode(macro) === "relative"
                          ? t("spreadsheet.macroRun")
                          : t("spreadsheet.macroRunAbsolute")
                      }
                      className={`${dialogParts.listItem} flex min-w-0 flex-1 items-center gap-2`}
                      data-testid={`spreadsheet-automation-recorded-run-${macro.name}`}
                    >
                      <SpreadsheetInsertIcon name="recordActions" size={20} />
                      <span className="truncate">{macro.name}</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => deleteMacro(macro.name)}
                      title={t("spreadsheet.macroDelete")}
                      aria-label={`${t("spreadsheet.macroDelete")}: ${macro.name}`}
                      className={dialogParts.iconButton}
                      data-testid={`spreadsheet-automation-recorded-delete-${macro.name}`}
                    >
                      <Trash2 size={16} aria-hidden="true" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </section>
          <div className={dialogParts.message}>
            <strong>
              {t("spreadsheet.automate.officeScriptsGroup")}
            </strong>{" "}
            {t("spreadsheet.automate.officeScriptsUnavailable")}
          </div>
        </OfficeDialogShell>
  );
}
