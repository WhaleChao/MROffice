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
import { type SheetModel } from "./spreadsheetSheetModel";
import { type DataFormRecord, type DataFormShape } from "../../../../lib/spreadsheet/data-form";
import { type Dispatch, type SetStateAction } from "react";
import { type ResolvedSpreadsheetVisibleControl } from "../../../../lib/spreadsheet/visible-command-authority";
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

/** What renderSpreadsheetDataFormDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetDataFormDialogInputs {
  closeDataForm: () => void;
  dataFormCriteria: string[] | null;
  dataFormDialogRef: RefObject<HTMLSpanElement | null>;
  dataFormRecord: DataFormRecord | null;
  dataFormShape: DataFormShape | null;
  dataFormStatus: string | null;
  dataFormValues: string[];
  executeVisibleCommand: (controlId: string) => Promise<boolean>;
  gridActiveCellFocusTarget: () => HTMLElement | null;
  setDataFormCriteria: Dispatch<SetStateAction<string[] | null>>;
  setDataFormDraft: Dispatch<SetStateAction<string[]>>;
  setDataFormRecordBaseline: Dispatch<SetStateAction<readonly string[] | null>>;
  sheet: SheetModel;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  visibleCommandDomActionProps: (controlId: string) => { onClick: () => Promise<boolean>; "data-control-id": string; disabled: boolean; hidden: boolean; title: string | undefined; };
  visibleCommandState: (controlId: string) => ResolvedSpreadsheetVisibleControl;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetDataFormDialog({
  closeDataForm,
  dataFormCriteria,
  dataFormDialogRef,
  dataFormRecord,
  dataFormShape,
  dataFormStatus,
  dataFormValues,
  executeVisibleCommand,
  gridActiveCellFocusTarget,
  setDataFormCriteria,
  setDataFormDraft,
  setDataFormRecordBaseline,
  sheet,
  t,
  visibleCommandDomActionProps,
  visibleCommandState,
}: RenderSpreadsheetDataFormDialogInputs) {
  return (
        <OfficeDialogShell
          open
          // Excel titles its data form with the worksheet's name.
          title={sheet.name}
          description={t("spreadsheet.dataFormMenu")}
          testIdPrefix="spreadsheet-data-form"
          defaultActionTestId="spreadsheet-data-form-close"
          defaultActionLabel={t("common.close")}
          cancelActionLabel={t("common.close")}
          showCancelAction={false}
          // A form left with no list (its range went away) still closes.
          onDefaultAction={() =>
            !dataFormShape || !dataFormRecord
              ? closeDataForm()
              : void executeVisibleCommand("XLSCTL.DATA.FORM.MRAI.DISMISS")
          }
          onCancel={() =>
            !dataFormShape || !dataFormRecord
              ? closeDataForm()
              : void executeVisibleCommand("XLSCTL.DATA.FORM.MRAI.DISMISS")
          }
          defaultActionDisabled={
            Boolean(dataFormShape && dataFormRecord) &&
            !visibleCommandState("XLSCTL.DATA.FORM.MRAI.DISMISS").enabled
          }
          defaultActionData={{
            "data-control-id": "XLSCTL.DATA.FORM.MRAI.DISMISS",
          }}
          returnFocusTarget={gridActiveCellFocusTarget}
          className="sm:max-w-[480px] max-h-[calc(100dvh-32px)] overflow-y-auto"
        >
          <span
            ref={dataFormDialogRef}
            style={{ display: "flex", flexDirection: "column", gap: 10 }}
          >
            {!dataFormShape || !dataFormRecord ? (
              <span
                className={dialogParts.hint}
                data-testid="spreadsheet-data-form-empty"
              >
                {t("spreadsheet.dataFormNeedsList")}
              </span>
            ) : (
              <>
                <span
                  style={{
                    display: "grid",
                    gridTemplateColumns: "minmax(0, 1fr) auto",
                    gap: 12,
                    alignItems: "start",
                  }}
                >
                  <span
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      gap: 8,
                      minWidth: 0,
                    }}
                  >
                    {dataFormRecord.fields.map((field, offset) => (
                      <label
                        key={`${field.label}-${field.col}`}
                        className={dialogParts.field}
                      >
                        <span className={dialogParts.label}>
                          {field.label}
                          {field.readOnly
                            ? ` (${t("spreadsheet.dataFormComputed")})`
                            : ""}
                        </span>
                        <input
                          value={
                            dataFormCriteria
                              ? (dataFormCriteria[offset] ?? "")
                              : (dataFormValues[offset] ?? "")
                          }
                          readOnly={!dataFormCriteria && field.readOnly}
                          onChange={(event) => {
                            const next = event.target.value;
                            if (dataFormCriteria) {
                              setDataFormCriteria((current) => {
                                const draft = [...(current ?? [])];
                                draft[offset] = next;
                                return draft;
                              });
                              return;
                            }
                            setDataFormRecordBaseline(
                              (current) =>
                                current ??
                                dataFormRecord.fields.map(
                                  (candidate) => candidate.value,
                                ),
                            );
                            setDataFormDraft((current) => {
                              const draft = dataFormRecord.fields.map(
                                (candidate, index) =>
                                  current[index] ?? candidate.value,
                              );
                              draft[offset] = next;
                              return draft;
                            });
                          }}
                          onKeyDown={(event) => {
                            if (
                              event.key !== "Enter" ||
                              event.nativeEvent.isComposing
                            ) {
                              return;
                            }
                            // Enter never presses Close: it steps to the
                            // next record (criteria are searched with the
                            // Find buttons instead).
                            event.preventDefault();
                            if (!dataFormCriteria) {
                              void executeVisibleCommand(
                                "XLSCTL.DATA.FORM.MRAI.NEXT",
                              );
                            }
                          }}
                          aria-label={field.label}
                          data-testid={`spreadsheet-data-form-field-${offset}`}
                        />
                      </label>
                    ))}
                  </span>
                  {/* Excel's order: the record count, then New, Delete,
                      Restore, Find Prev, Find Next and Criteria. */}
                  <span
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      gap: 6,
                      minWidth: 96,
                    }}
                  >
                    <span
                      className={dialogParts.hint}
                      data-testid="spreadsheet-data-form-position"
                    >
                      {dataFormCriteria
                        ? t("spreadsheet.dataFormCriteriaMode")
                        : `${dataFormRecord.index + 1} / ${dataFormShape.recordCount}`}
                    </span>
                    <button
                      type="button"
                      {...visibleCommandDomActionProps(
                        "XLSCTL.DATA.FORM.MRAI.NEW",
                      )}
                      data-testid="spreadsheet-data-form-new"
                      className={dialogParts.button}
                    >
                      <SpreadsheetDialogIcon name="add" />
                      {t("spreadsheet.dataFormNew")}
                    </button>
                    <button
                      type="button"
                      {...visibleCommandDomActionProps(
                        "XLSCTL.DATA.FORM.MRAI.DELETE",
                      )}
                      data-testid="spreadsheet-data-form-delete"
                      className={dialogParts.button}
                    >
                      <SpreadsheetDialogIcon name="delete" />
                      {t("presentation.comments.delete")}
                    </button>
                    <button
                      type="button"
                      {...visibleCommandDomActionProps(
                        "XLSCTL.DATA.FORM.MRAI.RESTORE",
                      )}
                      data-testid="spreadsheet-data-form-restore"
                      className={dialogParts.button}
                    >
                      <SpreadsheetDialogIcon name="reset" />
                      {t("window.restore")}
                    </button>
                    <button
                      type="button"
                      {...visibleCommandDomActionProps(
                        "XLSCTL.DATA.FORM.MRAI.PREVIOUS",
                      )}
                      data-testid="spreadsheet-data-form-prev"
                      className={dialogParts.button}
                    >
                      <SpreadsheetDialogIcon name="moveUp" />
                      {t("spreadsheet.dataFormPrev")}
                    </button>
                    <button
                      type="button"
                      {...visibleCommandDomActionProps(
                        "XLSCTL.DATA.FORM.MRAI.NEXT",
                      )}
                      data-testid="spreadsheet-data-form-next"
                      className={dialogParts.button}
                    >
                      <SpreadsheetDialogIcon name="moveDown" />
                      {t("spreadsheet.dataFormNext")}
                    </button>
                    <button
                      type="button"
                      {...visibleCommandDomActionProps(
                        "XLSCTL.DATA.FORM.MRAI.CRITERIA",
                      )}
                      data-active={dataFormCriteria ? "true" : "false"}
                      data-testid="spreadsheet-data-form-criteria"
                      className={dialogParts.button}
                    >
                      <SpreadsheetDialogIcon name="criteria" />
                      {t("spreadsheet.dataFormCriteria")}
                    </button>
                  </span>
                </span>
                {dataFormStatus ? (
                  <span
                    className={dialogParts.hint}
                    data-testid="spreadsheet-data-form-status"
                  >
                    {dataFormStatus}
                  </span>
                ) : null}
              </>
            )}
          </span>
        </OfficeDialogShell>
  );
}
