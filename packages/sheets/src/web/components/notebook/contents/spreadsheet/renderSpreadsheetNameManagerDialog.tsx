"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Fragment, type RefObject } from "react";
import { ChevronDown } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { type SpreadsheetVisibleCommandReason } from "../../../../lib/spreadsheet/visible-command-authority";
import { isValidName, isValidRef, MAX_NAME_COMMENT_LENGTH, type NamedRange } from "@mrnote/core/web/lib/spreadsheet/named-ranges.ts";
import { NAME_MANAGER_NO_FILTER, type NameManagerFilterCommand, isNameManagerFilterOn, toggleNameManagerFilter, type NameManagerFilter } from "../../../../lib/spreadsheet/name-manager-view";
import { DropdownMenu, DropdownMenuTrigger, DropdownMenuItem as BaseDropdownMenuItem, DropdownMenuSeparator, DropdownMenuCheckboxItem as BaseDropdownMenuCheckboxItem } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { type SheetModel } from "./spreadsheetSheetModel";
import { type Dispatch, type SetStateAction } from "react";
import { type ResolvedSpreadsheetVisibleControl } from "../../../../lib/spreadsheet/visible-command-authority";
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import { type SpreadsheetRangePick } from "./useSpreadsheetRangePick";
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

/** What renderSpreadsheetNameManagerDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetNameManagerDialogInputs {
  rangePick: SpreadsheetRangePick;
  beginCreateNameManagerEntry: () => void;
  beginEditNameManagerEntry: (entry: NamedRange) => void;
  closeNameManager: () => void;
  executeNameManagerApply: () => Promise<boolean>;
  executeNameManagerDelete: () => Promise<boolean>;
  formulaNameMutationCommandRef: RefObject<{ controlId: "XLSCTL.FORMULAS.NAME.DEFINE.APPLY"; name: string; ref: string; scope: "workbook" | "worksheet"; workbook: readonly SheetModel[]; sheet: SheetModel; activeSheetIndex: number; } | { controlId: "XLSCTL.FORMULAS.NAME.DELETE.APPLY"; entry: NamedRange; workbook: readonly SheetModel[]; sheet: SheetModel; activeSheetIndex: number; } | { controlId: "XLSCTL.FORMULAS.NAME.MANAGER.APPLY"; action: "create" | "edit"; source: NamedRange | null; name: string; ref: string; scope: "workbook" | "worksheet"; comment: string; workbook: readonly SheetModel[]; sheet: SheetModel; activeSheetIndex: number; } | { controlId: "XLSCTL.FORMULAS.NAME.MANAGER.DELETE"; source: NamedRange; workbook: readonly SheetModel[]; sheet: SheetModel; activeSheetIndex: number; } | null>;
  gridActiveCellFocusTarget: () => HTMLElement | null;
  nameManagerDraft: { name: string; ref: string; scope: "workbook" | "worksheet"; comment: string; };
  nameManagerDraftProblem: { field: "name" | "ref"; messageKey: string; } | null;
  nameManagerError: "spreadsheet.names.sourceChanged" | "spreadsheet.names.sourceReadOnly" | "spreadsheet.names.invalidUnique" | "spreadsheet.names.relativeRef" | "spreadsheet.names.renameCaptured" | "spreadsheet.names.renameCapturedByFormula" | null;
  nameManagerFilter: NameManagerFilter;
  nameManagerInputRef: RefObject<HTMLInputElement | null>;
  nameManagerMode: "edit" | "list" | "create";
  nameManagerOpen: boolean;
  nameManagerProblemId: string;
  nameManagerReturnFocusRef: RefObject<HTMLElement | null>;
  nameManagerRows: { readonly identity: string; readonly testId: string; readonly kind: "defined" | "table"; readonly entry: NamedRange | null; readonly name: string; readonly value: string; readonly refersTo: string; readonly scope: "workbook" | "worksheet"; readonly scopeLabel: string; readonly comment: string; readonly hasError: boolean; }[];
  nameManagerSelectedIdentity: string | null;
  nameManagerSort: "asc" | "desc";
  selectedNameManagerEntry: NamedRange | null;
  setNameManagerDraft: Dispatch<SetStateAction<{ name: string; ref: string; scope: "workbook" | "worksheet"; comment: string; }>>;
  setNameManagerError: Dispatch<SetStateAction<"spreadsheet.names.sourceChanged" | "spreadsheet.names.sourceReadOnly" | "spreadsheet.names.invalidUnique" | "spreadsheet.names.relativeRef" | "spreadsheet.names.renameCaptured" | "spreadsheet.names.renameCapturedByFormula" | null>>;
  setNameManagerFilter: Dispatch<SetStateAction<NameManagerFilter>>;
  setNameManagerMode: Dispatch<SetStateAction<"edit" | "list" | "create">>;
  setNameManagerSelectedIdentity: Dispatch<SetStateAction<string | null>>;
  setNameManagerSort: Dispatch<SetStateAction<"asc" | "desc">>;
  sheet: SheetModel;
  sheets: SheetModel[];
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  visibleCommandDomStateProps: (controlId: string) => { "data-control-id": string; disabled: boolean; hidden: boolean; title: string | undefined; };
  visibleCommandReason: (reasonKey: SpreadsheetVisibleCommandReason | null) => string | undefined;
  visibleCommandState: (controlId: string) => ResolvedSpreadsheetVisibleControl;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetNameManagerDialog({
  rangePick,
  beginCreateNameManagerEntry,
  beginEditNameManagerEntry,
  closeNameManager,
  executeNameManagerApply,
  executeNameManagerDelete,
  formulaNameMutationCommandRef,
  gridActiveCellFocusTarget,
  nameManagerDraft,
  nameManagerDraftProblem,
  nameManagerError,
  nameManagerFilter,
  nameManagerInputRef,
  nameManagerMode,
  nameManagerOpen,
  nameManagerProblemId,
  nameManagerReturnFocusRef,
  nameManagerRows,
  nameManagerSelectedIdentity,
  nameManagerSort,
  selectedNameManagerEntry,
  setNameManagerDraft,
  setNameManagerError,
  setNameManagerFilter,
  setNameManagerMode,
  setNameManagerSelectedIdentity,
  setNameManagerSort,
  sheet,
  sheets,
  t,
  visibleCommandDomStateProps,
  visibleCommandReason,
  visibleCommandState,
}: RenderSpreadsheetNameManagerDialogInputs) {
  // Collapsed while a range of it is picked on the sheet.
  if (rangePick.picking("spreadsheet-name-manager")) return null;
  return (
      <OfficeDialogShell
        open={nameManagerOpen}
        title={t("spreadsheet.formulas.nameManager")}
        description={
          t("spreadsheet.names.description")
        }
        defaultActionLabel={
          nameManagerMode === "list"
            ? t("pages.createShort")
            : t("common.save")
        }
        cancelActionLabel={
          nameManagerMode === "list"
            ? t("common.close")
            : t("library.back")
        }
        defaultActionDisabled={
          nameManagerMode === "list"
            ? false
            : !visibleCommandState(
                "XLSCTL.FORMULAS.NAME.MANAGER.APPLY",
              ).enabled ||
              !isValidName(nameManagerDraft.name.trim()) ||
              !isValidRef(
                nameManagerDraft.ref.trim().startsWith("=")
                  ? nameManagerDraft.ref.trim().slice(1).trim()
                  : nameManagerDraft.ref.trim(),
              ) ||
              nameManagerDraftProblem !== null ||
              (nameManagerMode === "edit" &&
                (!selectedNameManagerEntry ||
                  Boolean(selectedNameManagerEntry.nativeBlocker)))
        }
        onDefaultAction={() => {
          if (nameManagerMode === "list") {
            beginCreateNameManagerEntry();
          } else {
            void executeNameManagerApply();
          }
        }}
        onCancel={() => {
          if (nameManagerMode === "list") closeNameManager();
          else {
            formulaNameMutationCommandRef.current = null;
            setNameManagerMode("list");
            setNameManagerError(null);
          }
        }}
        initialFocusRef={
          nameManagerMode === "list" ? undefined : nameManagerInputRef
        }
        returnFocusRef={nameManagerReturnFocusRef}
        // Excel closes Name Manager back onto the worksheet, not onto the
        // command that opened it: the keyboard belongs to the active cell the
        // moment the dialog is gone, so arrows move the selection and Ctrl+Z
        // undoes a name change without a click first. The trigger button stays
        // the fallback when the grid has no live cell editor to take focus.
        returnFocusTarget={gridActiveCellFocusTarget}
        testIdPrefix="spreadsheet-name-manager"
        className="sm:max-w-[900px]"
      >
        {nameManagerMode === "list" ? (
          <>
            <div className="flex flex-wrap items-end justify-end gap-2">
              <button
                type="button"
                onClick={() =>
                  setNameManagerSort((current) =>
                    current === "asc" ? "desc" : "asc",
                  )
                }
                aria-label={
                  nameManagerSort === "asc"
                    ? t("spreadsheet.names.sortDescending")
                    : t("spreadsheet.names.sortAscending")
                }
                data-testid="spreadsheet-name-manager-sort"
                className={dialogParts.button}
              >
                <SpreadsheetDialogIcon
                  name={nameManagerSort === "asc" ? "sortAscending" : "sortDescending"}
                />
                {t("presentation.master.name")}
              </button>
              {/* Excel's Filter menu (XLS-A066): each command toggles, the
                  pairs exclude each other, Clear Filter turns them all off. */}
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <button
                    type="button"
                    data-testid="spreadsheet-name-manager-filter"
                    data-filtered={
                      nameManagerFilter.scope !== "all" ||
                      nameManagerFilter.errors !== "all" ||
                      nameManagerFilter.kind !== "all"
                        ? "true"
                        : undefined
                    }
                    className={dialogParts.button}
                  >
                    {t("spreadsheet.names.filter")}
                    <ChevronDown size={12} aria-hidden="true" />
                  </button>
                </DropdownMenuTrigger>
                <RibbonDropdownMenuContent
                  align="end"
                  data-testid="spreadsheet-name-manager-filter-menu"
                >
                  <BaseDropdownMenuItem
                    onSelect={() => setNameManagerFilter(NAME_MANAGER_NO_FILTER)}
                    data-testid="spreadsheet-name-manager-filter-clear"
                  >
                    {t("spreadsheet.names.clearFilter")}
                  </BaseDropdownMenuItem>
                  {(
                    [
                      [
                        ["worksheet", "spreadsheet.names.scopedToWorksheet"],
                        ["workbook", "spreadsheet.names.scopedToWorkbook"],
                      ],
                      [
                        ["with-errors", "spreadsheet.names.withErrors"],
                        ["without-errors", "spreadsheet.names.withoutErrors"],
                      ],
                      [
                        ["defined", "spreadsheet.names.filterDefinedNames"],
                        ["table", "spreadsheet.names.tableNames"],
                      ],
                    ] as const satisfies ReadonlyArray<
                      ReadonlyArray<readonly [NameManagerFilterCommand, string]>
                    >
                  ).map((pair, pairIndex) => (
                    <Fragment key={pairIndex}>
                      <DropdownMenuSeparator />
                      {pair.map(([command, labelKey]) => (
                        <BaseDropdownMenuCheckboxItem
                          key={command}
                          checked={isNameManagerFilterOn(nameManagerFilter, command)}
                          onCheckedChange={() =>
                            setNameManagerFilter((current) =>
                              toggleNameManagerFilter(current, command),
                            )
                          }
                          data-testid={`spreadsheet-name-manager-filter-${command}`}
                        >
                          {t(labelKey)}
                        </BaseDropdownMenuCheckboxItem>
                      ))}
                    </Fragment>
                  ))}
                </RibbonDropdownMenuContent>
              </DropdownMenu>
            </div>
            <div
              role="grid"
              aria-label={t("spreadsheet.names.definedNames")}
              className={`${dialogParts.list} max-h-72`}
              data-testid="spreadsheet-name-manager-grid"
            >
              <div
                role="row"
                className={`${dialogParts.hint} grid grid-cols-[minmax(100px,1fr)_minmax(90px,1fr)_minmax(140px,1.5fr)_minmax(80px,0.8fr)_minmax(90px,1fr)] gap-3 px-2 py-1 font-semibold`}
              >
                <span role="columnheader">{t("presentation.master.name")}</span>
                <span role="columnheader">{t("spreadsheet.names.value")}</span>
                <span role="columnheader">
                  {t("spreadsheet.names.refersTo")}
                </span>
                <span role="columnheader">{t("spreadsheet.names.scope")}</span>
                <span role="columnheader">{t("spreadsheet.names.comment")}</span>
              </div>
              {nameManagerRows.length ? (
                nameManagerRows.map((row) => {
                  const selected = row.identity === nameManagerSelectedIdentity;
                  return (
                    <button
                      key={row.identity}
                      type="button"
                      role="row"
                      aria-selected={selected}
                      onClick={() => setNameManagerSelectedIdentity(row.identity)}
                      onDoubleClick={() => {
                        if (row.entry) beginEditNameManagerEntry(row.entry);
                      }}
                      data-testid={`spreadsheet-name-manager-row-${row.testId}`}
                      data-kind={row.kind}
                      className={`${dialogParts.listItem} grid w-full grid-cols-[minmax(100px,1fr)_minmax(90px,1fr)_minmax(140px,1.5fr)_minmax(80px,0.8fr)_minmax(90px,1fr)] gap-3`}
                    >
                      <span role="gridcell" className="min-w-0 truncate font-medium">
                        {row.name}
                        {row.entry?.nativeBlocker ? (
                          <span className={`${dialogParts.message} ps-2`} data-tone="warning">
                            {t("ai.readOnlyLock")}
                          </span>
                        ) : null}
                      </span>
                      <span
                        role="gridcell"
                        className="min-w-0 truncate"
                        title={row.value}
                        data-testid={`spreadsheet-name-manager-value-${row.testId}`}
                      >
                        {row.value}
                      </span>
                      <span role="gridcell" className="min-w-0 truncate font-mono" title={row.refersTo}>
                        {row.refersTo}
                      </span>
                      <span role="gridcell" className="min-w-0 truncate">
                        {row.scopeLabel}
                      </span>
                      <span
                        role="gridcell"
                        className="min-w-0 truncate"
                        title={row.comment || undefined}
                        data-testid={`spreadsheet-name-manager-comment-${row.testId}`}
                      >
                        {row.comment}
                      </span>
                    </button>
                  );
                })
              ) : (
                <p className={`${dialogParts.hint} px-3 py-6 text-center`}>
                  {t("spreadsheet.names.emptyFiltered")}
                </p>
              )}
            </div>
            {nameManagerSelectedIdentity?.startsWith("\u0001table") ? (
              <p
                role="status"
                data-testid="spreadsheet-name-manager-table-read-only"
                className={dialogParts.hint}
              >
                {t("spreadsheet.names.tableReadOnly")}
              </p>
            ) : null}
            {selectedNameManagerEntry?.nativeBlocker ? (
              <p
                role="status"
                data-testid="spreadsheet-name-manager-native-blocker"
                className={dialogParts.message}
                data-tone="warning"
              >
                {t("spreadsheet.names.nativeReadOnly")}
              </p>
            ) : null}
            <div className="flex flex-wrap gap-2">
              <button
                type="button"
                disabled={
                  !selectedNameManagerEntry ||
                  Boolean(selectedNameManagerEntry.nativeBlocker)
                }
                onClick={() =>
                  selectedNameManagerEntry &&
                  beginEditNameManagerEntry(selectedNameManagerEntry)
                }
                data-testid="spreadsheet-name-manager-edit"
                className={dialogParts.compactButton}
              >
                <SpreadsheetDialogIcon name="edit" />
                {t("memoryGraph.drawer.actions.edit")}
              </button>
              <button
                type="button"
                {...visibleCommandDomStateProps(
                  "XLSCTL.FORMULAS.NAME.MANAGER.DELETE",
                )}
                disabled={
                  !selectedNameManagerEntry ||
                  Boolean(selectedNameManagerEntry.nativeBlocker) ||
                  !visibleCommandState(
                    "XLSCTL.FORMULAS.NAME.MANAGER.DELETE",
                  ).enabled
                }
                onClick={() => void executeNameManagerDelete()}
                data-testid="spreadsheet-name-manager-delete"
                className={dialogParts.compactButton}
              >
                <SpreadsheetDialogIcon name="delete" />
                {t("spreadsheet.nameDelete")}
              </button>
            </div>
          </>
        ) : (
          <>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("presentation.master.name")}</span>
              <input
                ref={nameManagerInputRef}
                value={nameManagerDraft.name}
                maxLength={255}
                onChange={(event) => {
                  setNameManagerDraft((current) => ({
                    ...current,
                    name: event.target.value,
                  }));
                  setNameManagerError(null);
                }}
                aria-invalid={
                  nameManagerDraftProblem?.field === "name" ? true : undefined
                }
                aria-describedby={
                  nameManagerDraftProblem?.field === "name"
                    ? `${nameManagerProblemId}-name`
                    : undefined
                }
                data-testid="spreadsheet-name-manager-name"
              />
              {nameManagerDraftProblem?.field === "name" ? (
                <span
                  id={`${nameManagerProblemId}-name`}
                  aria-live="polite"
                  data-testid="spreadsheet-name-manager-name-problem"
                  className={dialogParts.error}
                >
                  {t(nameManagerDraftProblem.messageKey)}
                </span>
              ) : null}
            </label>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.names.scope")}</span>
              <select
                value={nameManagerDraft.scope}
                disabled={nameManagerMode === "edit"}
                onChange={(event) =>
                  setNameManagerDraft((current) => ({
                    ...current,
                    scope:
                      event.target.value === "worksheet"
                        ? "worksheet"
                        : "workbook",
                  }))
                }
                data-testid="spreadsheet-name-manager-scope"
              >
                <option value="workbook">
                  {t("spreadsheet.help.workbook")}
                </option>
                <option value="worksheet">
                  {nameManagerMode === "edit" &&
                  selectedNameManagerEntry?.scope === "worksheet"
                    ? sheets.find(
                        (candidate) =>
                          candidate.id === selectedNameManagerEntry.scopeSheetId,
                      )?.name ?? sheet.name
                    : sheet.name}
                </option>
              </select>
            </label>
            {/* Excel's New / Edit Name box: Comment sits between Scope and
                Refers to, up to 255 characters (XLS-A066). */}
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.names.comment")}</span>
              <textarea
                value={nameManagerDraft.comment}
                maxLength={MAX_NAME_COMMENT_LENGTH}
                rows={3}
                onChange={(event) => {
                  setNameManagerDraft((current) => ({
                    ...current,
                    comment: event.target.value,
                  }));
                  setNameManagerError(null);
                }}
                data-testid="spreadsheet-name-manager-comment"
                style={{ resize: "vertical" }}
              />
            </label>
            <label className={dialogParts.field}>
              <span className={dialogParts.label}>{t("spreadsheet.names.refersTo")}</span>
              <span style={{ display: "flex", gap: 4 }}>
                <input
                  value={nameManagerDraft.ref}
                  maxLength={8192}
                  onChange={(event) => {
                    setNameManagerDraft((current) => ({
                      ...current,
                      ref: event.target.value,
                    }));
                    setNameManagerError(null);
                  }}
                  aria-invalid={
                    nameManagerDraftProblem?.field === "ref" ? true : undefined
                  }
                  aria-describedby={
                    nameManagerDraftProblem?.field === "ref"
                      ? `${nameManagerProblemId}-ref`
                      : undefined
                  }
                  data-testid="spreadsheet-name-manager-ref"
                  className={dialogParts.mono}
                  style={{ flex: 1, minWidth: 0 }}
                />
                <SpreadsheetCollapseDialogButton
                  label={t("spreadsheet.cfDialog.collapse")}
                  onClick={() =>
                    rangePick.start({
                      id: "spreadsheet-name-manager-ref",
                      dialogLabel: t("spreadsheet.formulas.nameManager"),
                      inputLabel: t("spreadsheet.names.refersTo"),
                      style: "sheetFormula",
                      otherSheets: true,
                      value: String(nameManagerDraft.ref),
                      onChange: (text: string) => {
                        setNameManagerDraft((current) => ({
                          ...current,
                          ref: text,
                        }));
                        setNameManagerError(null);
                      },
                    })
                  }
                  testId="spreadsheet-name-manager-ref-pick"
                />
              </span>
              {nameManagerDraftProblem?.field === "ref" ? (
                <span
                  id={`${nameManagerProblemId}-ref`}
                  aria-live="polite"
                  data-testid="spreadsheet-name-manager-ref-problem"
                  className={dialogParts.error}
                >
                  {t(nameManagerDraftProblem.messageKey)}
                </span>
              ) : null}
            </label>
            {nameManagerError ? (
              <p
                role="alert"
                data-testid="spreadsheet-name-manager-error"
                className={dialogParts.error}
              >
                {t(nameManagerError)}
              </p>
            ) : null}
            {!visibleCommandState(
              "XLSCTL.FORMULAS.NAME.MANAGER.APPLY",
            ).enabled ? (
              <p className={dialogParts.message} data-tone="warning">
                {visibleCommandReason(
                  visibleCommandState(
                    "XLSCTL.FORMULAS.NAME.MANAGER.APPLY",
                  ).reasonKey,
                )}
              </p>
            ) : null}
          </>
        )}
      </OfficeDialogShell>
  );
}
