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
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import { spreadsheetDataValidationPatchbackRuleBlocker, type SpreadsheetVisibleCommandReason } from "../../../../lib/spreadsheet/visible-command-authority";
import { type ValidationOperator, validationOperatorNeedsSecondValue, type ErrorStyle, type ValidationRule } from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";
import { type ValidationDialogKind, validationBoundLabelKeys } from "./spreadsheetDialogDrafts";
import { dependentSourceAddress } from "../../../../lib/spreadsheet/validation-buttons";
import { type Dispatch, type SetStateAction } from "react";
import { type ResolvedSpreadsheetVisibleControl, type SpreadsheetDataValidationPatchbackBlocker } from "../../../../lib/spreadsheet/visible-command-authority";
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

/** What renderSpreadsheetDataValidationDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetDataValidationDialogInputs {
  closeValidationDialog: () => void;
  executeDataValidationApply: () => Promise<boolean>;
  gridActiveCellFocusTarget: () => HTMLElement | null;
  setValApplyToSame: Dispatch<SetStateAction<boolean>>;
  setValDateMax: Dispatch<SetStateAction<string>>;
  setValDateMin: Dispatch<SetStateAction<string>>;
  setValDepDefault: Dispatch<SetStateAction<string>>;
  setValDepOptions: Dispatch<SetStateAction<string>>;
  setValDepPerRow: Dispatch<SetStateAction<boolean>>;
  setValDepSourceCol: Dispatch<SetStateAction<string>>;
  setValDepSourceRow: Dispatch<SetStateAction<string>>;
  setValErrorMessage: Dispatch<SetStateAction<string>>;
  setValErrorStyle: Dispatch<SetStateAction<ErrorStyle>>;
  setValErrorTitle: Dispatch<SetStateAction<string>>;
  setValFormula: Dispatch<SetStateAction<string>>;
  setValidationClearPending: Dispatch<SetStateAction<boolean>>;
  setValidationDialogTab: Dispatch<SetStateAction<"error" | "settings" | "input" | "ime">>;
  setValidationError: Dispatch<SetStateAction<string | null>>;
  setValIgnoreBlank: Dispatch<SetStateAction<boolean>>;
  setValInputMessage: Dispatch<SetStateAction<string>>;
  setValInputTitle: Dispatch<SetStateAction<string>>;
  setValInteger: Dispatch<SetStateAction<boolean>>;
  setValKind: Dispatch<SetStateAction<ValidationDialogKind>>;
  setValListValues: Dispatch<SetStateAction<string>>;
  setValMax: Dispatch<SetStateAction<string>>;
  setValMaxLength: Dispatch<SetStateAction<string>>;
  setValMin: Dispatch<SetStateAction<string>>;
  setValMinLength: Dispatch<SetStateAction<string>>;
  setValNotAfterToday: Dispatch<SetStateAction<boolean>>;
  setValNotBeforeToday: Dispatch<SetStateAction<boolean>>;
  setValOperator: Dispatch<SetStateAction<ValidationOperator>>;
  setValShowDropDown: Dispatch<SetStateAction<boolean>>;
  setValShowErrorMessage: Dispatch<SetStateAction<boolean>>;
  setValShowInputMessage: Dispatch<SetStateAction<boolean>>;
  setValTimeMax: Dispatch<SetStateAction<string>>;
  setValTimeMin: Dispatch<SetStateAction<string>>;
  setValWeekdaysOnly: Dispatch<SetStateAction<boolean>>;
  spreadsheetNativeWriteRoute: "patchback" | "from-scratch";
  startValidationSourcePick: () => void;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  valApplyToSame: boolean;
  valDateMax: string;
  valDateMin: string;
  valDepDefault: string;
  valDepOptions: string;
  valDepPerRow: boolean;
  valDepSourceCol: string;
  valDepSourceRow: string;
  valErrorMessage: string;
  valErrorStyle: ErrorStyle;
  valErrorTitle: string;
  valFormula: string;
  validationClearPending: boolean;
  validationDialogTab: "error" | "settings" | "input" | "ime";
  validationDialogTabs: readonly [{ readonly id: "settings"; readonly label: string; }, { readonly id: "input"; readonly label: string; }, { readonly id: "error"; readonly label: string; }, { readonly id: "ime"; readonly label: string; }];
  validationDraftRule: ValidationRule | null;
  validationError: string | null;
  validationImeUnavailableReason: string;
  validationKindRef: RefObject<HTMLSelectElement | null>;
  validationOpenedRule: ValidationRule | null;
  validationPatchbackBlockerMessage: (blocker: SpreadsheetDataValidationPatchbackBlocker) => string;
  valIgnoreBlank: boolean;
  valInputMessage: string;
  valInputTitle: string;
  valInteger: boolean;
  valKind: ValidationDialogKind;
  valListValues: string;
  valMax: string;
  valMaxLength: string;
  valMin: string;
  valMinLength: string;
  valNotAfterToday: boolean;
  valNotBeforeToday: boolean;
  valOperator: ValidationOperator;
  valShowDropDown: boolean;
  valShowErrorMessage: boolean;
  valShowInputMessage: boolean;
  valTimeMax: string;
  valTimeMin: string;
  valWeekdaysOnly: boolean;
  visibleCommandDomActionProps: (controlId: string) => { onClick: () => Promise<boolean>; "data-control-id": string; disabled: boolean; hidden: boolean; title: string | undefined; };
  visibleCommandReason: (reasonKey: SpreadsheetVisibleCommandReason | null) => string | undefined;
  visibleCommandState: (controlId: string) => ResolvedSpreadsheetVisibleControl;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetDataValidationDialog({
  closeValidationDialog,
  executeDataValidationApply,
  gridActiveCellFocusTarget,
  setValApplyToSame,
  setValDateMax,
  setValDateMin,
  setValDepDefault,
  setValDepOptions,
  setValDepPerRow,
  setValDepSourceCol,
  setValDepSourceRow,
  setValErrorMessage,
  setValErrorStyle,
  setValErrorTitle,
  setValFormula,
  setValidationClearPending,
  setValidationDialogTab,
  setValidationError,
  setValIgnoreBlank,
  setValInputMessage,
  setValInputTitle,
  setValInteger,
  setValKind,
  setValListValues,
  setValMax,
  setValMaxLength,
  setValMin,
  setValMinLength,
  setValNotAfterToday,
  setValNotBeforeToday,
  setValOperator,
  setValShowDropDown,
  setValShowErrorMessage,
  setValShowInputMessage,
  setValTimeMax,
  setValTimeMin,
  setValWeekdaysOnly,
  spreadsheetNativeWriteRoute,
  startValidationSourcePick,
  t,
  valApplyToSame,
  valDateMax,
  valDateMin,
  valDepDefault,
  valDepOptions,
  valDepPerRow,
  valDepSourceCol,
  valDepSourceRow,
  valErrorMessage,
  valErrorStyle,
  valErrorTitle,
  valFormula,
  validationClearPending,
  validationDialogTab,
  validationDialogTabs,
  validationDraftRule,
  validationError,
  validationImeUnavailableReason,
  validationKindRef,
  validationOpenedRule,
  validationPatchbackBlockerMessage,
  valIgnoreBlank,
  valInputMessage,
  valInputTitle,
  valInteger,
  valKind,
  valListValues,
  valMax,
  valMaxLength,
  valMin,
  valMinLength,
  valNotAfterToday,
  valNotBeforeToday,
  valOperator,
  valShowDropDown,
  valShowErrorMessage,
  valShowInputMessage,
  valTimeMax,
  valTimeMin,
  valWeekdaysOnly,
  visibleCommandDomActionProps,
  visibleCommandReason,
  visibleCommandState,
}: RenderSpreadsheetDataValidationDialogInputs) {
  return (
        <OfficeDialogShell
          open
          title={t("spreadsheet.dataValidationMenu")}
          description={null}
          testIdPrefix="spreadsheet-validation"
          defaultActionTestId="spreadsheet-validation-apply"
          defaultActionLabel={t("spreadsheet.validationApply")}
          cancelActionLabel={t("spreadsheet.validationCancel")}
          onDefaultAction={() => void executeDataValidationApply()}
          onCancel={closeValidationDialog}
          defaultActionDisabled={
            !visibleCommandState("XLSCTL.DATA.VALIDATION.MRAI.APPLY").enabled
          }
          defaultActionData={{
            "data-control-id": "XLSCTL.DATA.VALIDATION.MRAI.APPLY",
            "data-control-surface": "data-validation-dialog",
            "data-control-occurrence-id": "apply",
          }}
          initialFocusRef={validationKindRef}
          returnFocusTarget={gridActiveCellFocusTarget}
          className="sm:max-w-[520px] max-h-[calc(100dvh-32px)] overflow-y-auto"
          tabStrip={
            <span
              role="tablist"
              aria-label={t("spreadsheet.dataValidationMenu")}
              data-testid="spreadsheet-validation-tabs"
              className={dialogParts.tabs}
            >
              {validationDialogTabs.map((tab) => {
                const disabled = tab.id === "ime";
                return (
                  <button
                    key={tab.id}
                    id={`spreadsheet-validation-tab-${tab.id}`}
                    type="button"
                    role="tab"
                    aria-selected={validationDialogTab === tab.id}
                    aria-controls={`spreadsheet-validation-panel-${tab.id}`}
                    aria-disabled={disabled}
                    tabIndex={validationDialogTab === tab.id ? 0 : -1}
                    aria-describedby={
                      disabled ? "spreadsheet-validation-ime-reason" : undefined
                    }
                    disabled={disabled}
                    title={disabled ? validationImeUnavailableReason : undefined}
                    onClick={() => {
                      setValidationDialogTab(tab.id);
                      setValidationError(null);
                    }}
                    onKeyDown={(event) => {
                      if (
                        event.key !== "ArrowLeft" &&
                        event.key !== "ArrowRight" &&
                        event.key !== "Home" &&
                        event.key !== "End"
                      ) {
                        return;
                      }
                      event.preventDefault();
                      const enabledTabs = ["settings", "input", "error"] as const;
                      const current = Math.max(
                        0,
                        enabledTabs.indexOf(
                          validationDialogTab as (typeof enabledTabs)[number],
                        ),
                      );
                      const next =
                        event.key === "Home"
                          ? enabledTabs[0]
                          : event.key === "End"
                            ? enabledTabs[enabledTabs.length - 1]
                            : enabledTabs[
                                (current +
                                  (event.key === "ArrowRight" ? 1 : -1) +
                                  enabledTabs.length) %
                                  enabledTabs.length
                              ];
                      setValidationDialogTab(next);
                      setValidationError(null);
                      window.requestAnimationFrame(() =>
                        document
                          .getElementById(`spreadsheet-validation-tab-${next}`)
                          ?.focus({ preventScroll: true }),
                      );
                    }}
                    data-testid={`spreadsheet-validation-tab-${tab.id}`}
                    className={dialogParts.tab}
                  >
                    {tab.label}
                  </button>
                );
              })}
            </span>
          }
          footerStart={
            <button
              type="button"
              {...visibleCommandDomActionProps(
                "XLSCTL.DATA.VALIDATION.MRAI.CLEAR",
              )}
              data-control-surface="data-validation-dialog"
              data-control-occurrence-id="clear-draft"
              aria-pressed={validationClearPending}
              data-testid="spreadsheet-validation-clear"
              className={dialogParts.button}
            >
              <SpreadsheetDialogIcon name="clear" />
              {t("spreadsheet.validationClear")}
            </button>
          }
        >
          <span
            data-clear-pending={validationClearPending ? "true" : "false"}
            onChangeCapture={() => {
              if (validationClearPending) setValidationClearPending(false);
            }}
            style={{ display: "contents" }}
          >
            <span
              id="spreadsheet-validation-ime-reason"
              style={{
                position: "absolute",
                width: 1,
                height: 1,
                overflow: "hidden",
                clip: "rect(0 0 0 0)",
              }}
            >
              {validationImeUnavailableReason}
            </span>
            <span
              id="spreadsheet-validation-panel-settings"
              role="tabpanel"
              aria-labelledby="spreadsheet-validation-tab-settings"
              data-testid="spreadsheet-validation-panel-settings"
              style={{
                display: validationDialogTab === "settings" ? "flex" : "none",
                flexDirection: "column",
                gap: 6,
              }}
            >
              {/* Excel's Settings tab: Validation criteria (Allow, Ignore blank,
                  Data and the boxes that follow them), then "Apply these
                  changes to all other cells with the same settings". */}
              <span
                className={dialogParts.sectionTitle}
              >
                {t("spreadsheet.validationCriteria")}
              </span>
              <span style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <label
                  className={dialogParts.row}
                  style={{ flex: 1 }}
                >
                  {t("spreadsheet.validationKind")}
                  <select
                    value={
                      valKind === "number"
                        ? valInteger
                          ? "whole"
                          : "number"
                        : valKind
                    }
                    onChange={(event) => {
                      const next = event.target.value;
                      setValKind(
                        next === "whole"
                          ? "number"
                          : (next as ValidationDialogKind),
                      );
                      setValInteger(next === "whole");
                      setValOperator("between");
                      setValidationError(null);
                    }}
                    data-testid="spreadsheet-validation-kind"
                    ref={validationKindRef}
                    style={{ flex: 1 }}
                  >
                    <option value="any">{t("spreadsheet.validationKindAny")}</option>
                    <option value="whole">
                      {t("spreadsheet.validationKindWhole")}
                    </option>
                    <option value="number">
                      {t("spreadsheet.validationKindNumber")}
                    </option>
                    <option value="list">
                      {t("spreadsheet.validationKindList")}
                    </option>
                    <option value="date">
                      {t("spreadsheet.validationKindDate")}
                    </option>
                    <option value="time">
                      {t("spreadsheet.validationKindTime")}
                    </option>
                    <option value="text">
                      {t("spreadsheet.validationKindText")}
                    </option>
                    <option value="formula">
                      {t("spreadsheet.validationKindFormula")}
                    </option>
                    <option value="dependent">
                      {t("spreadsheet.validationKindDependent")}
                    </option>
                  </select>
                </label>
                <span style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  <label
                    className={dialogParts.check} style={{ opacity: valKind === "any" ? 0.5 : 1 }}
                  >
                    <input
                      type="checkbox"
                      checked={valIgnoreBlank}
                      disabled={valKind === "any"}
                      onChange={(event) => setValIgnoreBlank(event.target.checked)}
                      data-testid="spreadsheet-validation-ignore-blank"
                    />
                    {t("spreadsheet.validationIgnoreBlank")}
                  </label>
                  {valKind === "list" ? (
                    <label
                      className={dialogParts.check}
                    >
                      <input
                        type="checkbox"
                        checked={valShowDropDown}
                        onChange={(event) =>
                          setValShowDropDown(event.target.checked)
                        }
                        data-testid="spreadsheet-validation-show-dropdown"
                      />
                      {t("spreadsheet.validationInCellDropdown")}
                    </label>
                  ) : null}
                </span>
              </span>
            {valKind === "number" ||
            valKind === "text" ||
            valKind === "date" ||
            valKind === "time" ? (
              <label
                className={dialogParts.row}
              >
                {t("spreadsheet.validationOperator")}
                <select
                  value={valOperator}
                  onChange={(event) => {
                    setValOperator(event.target.value as ValidationOperator);
                    setValidationError(null);
                  }}
                  data-testid="spreadsheet-validation-operator"
                  style={{ flex: 1 }}
                >
                  <option value="between">
                    {t("spreadsheet.validationOperatorBetween")}
                  </option>
                  <option value="notBetween">
                    {t("spreadsheet.validationOperatorNotBetween")}
                  </option>
                  <option value="equal">
                    {t("spreadsheet.validationOperatorEqual")}
                  </option>
                  <option value="notEqual">
                    {t("spreadsheet.validationOperatorNotEqual")}
                  </option>
                  <option value="greaterThan">
                    {t("spreadsheet.validationOperatorGreaterThan")}
                  </option>
                  <option value="lessThan">
                    {t("spreadsheet.validationOperatorLessThan")}
                  </option>
                  <option value="greaterThanOrEqual">
                    {t("spreadsheet.validationOperatorGreaterThanOrEqual")}
                  </option>
                  <option value="lessThanOrEqual">
                    {t("spreadsheet.validationOperatorLessThanOrEqual")}
                  </option>
                </select>
              </label>
            ) : null}
            {valKind === "number" ||
            valKind === "text" ||
            valKind === "date" ||
            valKind === "time" ? (
              // The captions over the bound boxes follow the Data operator.
              <span
                aria-hidden="true"
                data-testid="spreadsheet-validation-bound-labels"
                className={dialogParts.label}
                style={{ display: "flex", gap: 4 }}
              >
                <span style={{ flex: 1 }}>
                  {t(validationBoundLabelKeys(valKind, valOperator)[0])}
                </span>
                {validationOperatorNeedsSecondValue(valOperator) ? (
                  <span style={{ flex: 1 }}>
                    {t(validationBoundLabelKeys(valKind, valOperator)[1])}
                  </span>
                ) : null}
              </span>
            ) : null}
            {valKind === "list" ? (
              <label
                className={dialogParts.field}
              >
                {t("spreadsheet.validationSource")}
                <span style={{ display: "flex", gap: 4 }}>
                  <input
                    value={valListValues}
                    onChange={(event) => {
                      setValListValues(event.target.value);
                      if (validationError) setValidationError(null);
                    }}
                    placeholder={t("spreadsheet.validationListPlaceholder")}
                    data-testid="spreadsheet-validation-list-values"
                    style={{ flex: 1, minWidth: 0 }}
                  />
                  <SpreadsheetCollapseDialogButton
                    label={t("spreadsheet.cfDialog.collapse")}
                    onClick={startValidationSourcePick}
                    testId="spreadsheet-validation-source-pick"
                  />
                </span>
              </label>
            ) : null}
            {valKind === "number" ? (
              <span style={{ display: "flex", gap: 4 }}>
                <input
                  value={valMin}
                  onChange={(event) => {
                    setValMin(event.target.value);
                    if (validationError) setValidationError(null);
                  }}
                  aria-label={t(validationBoundLabelKeys(valKind, valOperator)[0])}
                  inputMode="decimal"
                  data-testid="spreadsheet-validation-min"
                  style={{ flex: 1, minWidth: 0 }}
                />
                {validationOperatorNeedsSecondValue(valOperator) ? (
                  <input
                    value={valMax}
                    onChange={(event) => {
                      setValMax(event.target.value);
                      if (validationError) setValidationError(null);
                    }}
                    aria-label={t(validationBoundLabelKeys(valKind, valOperator)[1])}
                    inputMode="decimal"
                    data-testid="spreadsheet-validation-max"
                    style={{ flex: 1, minWidth: 0 }}
                  />
                ) : null}
              </span>
            ) : null}
            {valKind === "text" ? (
              <span style={{ display: "flex", gap: 4 }}>
                <input
                  value={valMinLength}
                  onChange={(event) => {
                    setValMinLength(event.target.value);
                    if (validationError) setValidationError(null);
                  }}
                  aria-label={t(validationBoundLabelKeys(valKind, valOperator)[0])}
                  inputMode="numeric"
                  data-testid="spreadsheet-validation-min-length"
                  style={{ flex: 1, minWidth: 0 }}
                />
                {validationOperatorNeedsSecondValue(valOperator) ? (
                  <input
                    value={valMaxLength}
                    onChange={(event) => {
                      setValMaxLength(event.target.value);
                      if (validationError) setValidationError(null);
                    }}
                    aria-label={t(validationBoundLabelKeys(valKind, valOperator)[1])}
                    inputMode="numeric"
                    data-testid="spreadsheet-validation-max-length"
                    style={{ flex: 1, minWidth: 0 }}
                  />
                ) : null}
              </span>
            ) : null}
            {valKind === "formula" ? (
              <label
                className={dialogParts.field}
              >
                {t("spreadsheet.validationFormula")}
                <textarea
                  value={valFormula}
                  onChange={(event) => {
                    setValFormula(event.target.value);
                    if (validationError) setValidationError(null);
                  }}
                  placeholder={t("spreadsheet.validationFormulaPlaceholder")}
                  rows={2}
                  data-testid="spreadsheet-validation-formula"
                  style={{ resize: "vertical", fontFamily: "var(--font-mono, monospace)" }}
                />
              </label>
            ) : null}
            {valKind === "date" ? (
              <span style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <span style={{ display: "flex", gap: 4 }}>
                  <input
                    value={valDateMin}
                    onChange={(event) => {
                      setValDateMin(event.target.value);
                      if (validationError) setValidationError(null);
                    }}
                    aria-label={t(validationBoundLabelKeys(valKind, valOperator)[0])}
                    type="date"
                    data-testid="spreadsheet-validation-date-min"
                    style={{ flex: 1, minWidth: 0 }}
                  />
                  {validationOperatorNeedsSecondValue(valOperator) ? (
                    <input
                      value={valDateMax}
                      onChange={(event) => {
                        setValDateMax(event.target.value);
                        if (validationError) setValidationError(null);
                      }}
                      placeholder={t("spreadsheet.validationDateMax")}
                      aria-label={t("spreadsheet.validationDateMax")}
                      type="date"
                      data-testid="spreadsheet-validation-date-max"
                      style={{ flex: 1, minWidth: 0 }}
                    />
                  ) : null}
                </span>
                <label
                  className={dialogParts.check}
                >
                  <input
                    type="checkbox"
                    checked={valWeekdaysOnly}
                    onChange={(event) => setValWeekdaysOnly(event.target.checked)}
                    data-testid="spreadsheet-validation-date-weekdays"
                  />
                  {t("spreadsheet.validationDateWeekdaysOnly")}
                </label>
                <label
                  className={dialogParts.check}
                >
                  <input
                    type="checkbox"
                    checked={valNotBeforeToday}
                    onChange={(event) =>
                      setValNotBeforeToday(event.target.checked)
                    }
                    data-testid="spreadsheet-validation-date-not-before-today"
                  />
                  {t("spreadsheet.validationDateNotBeforeToday")}
                </label>
                <label
                  className={dialogParts.check}
                >
                  <input
                    type="checkbox"
                    checked={valNotAfterToday}
                    onChange={(event) =>
                      setValNotAfterToday(event.target.checked)
                    }
                    data-testid="spreadsheet-validation-date-not-after-today"
                  />
                  {t("spreadsheet.validationDateNotAfterToday")}
                </label>
              </span>
            ) : null}
            {valKind === "time" ? (
              <span style={{ display: "flex", gap: 4 }}>
                <input
                  value={valTimeMin}
                  onChange={(event) => {
                    setValTimeMin(event.target.value);
                    if (validationError) setValidationError(null);
                  }}
                  type="time"
                  step={1}
                  aria-label={t(validationBoundLabelKeys(valKind, valOperator)[0])}
                  data-testid="spreadsheet-validation-time-min"
                  style={{ flex: 1, minWidth: 0 }}
                />
                {validationOperatorNeedsSecondValue(valOperator) ? (
                  <input
                    value={valTimeMax}
                    onChange={(event) => {
                      setValTimeMax(event.target.value);
                      if (validationError) setValidationError(null);
                    }}
                    type="time"
                    step={1}
                    aria-label={t("spreadsheet.validationTimeMax")}
                    data-testid="spreadsheet-validation-time-max"
                    style={{ flex: 1, minWidth: 0 }}
                  />
                ) : null}
              </span>
            ) : null}
            {valKind === "dependent" ? (
              <span style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <span style={{ display: "flex", gap: 4 }}>
                  <input
                    value={valDepSourceRow}
                    onChange={(event) => {
                      setValDepSourceRow(event.target.value);
                      if (validationError) setValidationError(null);
                    }}
                    placeholder={t("spreadsheet.validationDependentSourceRow")}
                    inputMode="numeric"
                    data-testid="spreadsheet-validation-dep-source-row"
                    style={{ flex: 1, minWidth: 0 }}
                  />
                  <input
                    value={valDepSourceCol}
                    onChange={(event) => {
                      setValDepSourceCol(event.target.value);
                      if (validationError) setValidationError(null);
                    }}
                    placeholder={t("spreadsheet.validationDependentSourceCol")}
                    data-testid="spreadsheet-validation-dep-source-col"
                    style={{ flex: 1, minWidth: 0 }}
                  />
                  {/* The cell the two fields name, as the grid labels it
                      (row from 1, column by letters: live audit S-192). */}
                  <output
                    data-testid="spreadsheet-validation-dep-source-address"
                    style={{ alignSelf: "center", minWidth: "3.5em", fontVariantNumeric: "tabular-nums" }}
                  >
                    {dependentSourceAddress(valDepSourceRow, valDepSourceCol) ? `= ${dependentSourceAddress(valDepSourceRow, valDepSourceCol)}` : ""}
                  </output>
                </span>
                <label
                  className={dialogParts.check}
                >
                  <input
                    type="checkbox"
                    checked={valDepPerRow}
                    onChange={(event) => setValDepPerRow(event.target.checked)}
                    data-testid="spreadsheet-validation-dep-per-row"
                  />
                  {t("spreadsheet.validationDependentPerRow")}
                </label>
                <textarea
                  value={valDepOptions}
                  onChange={(event) => {
                    setValDepOptions(event.target.value);
                    if (validationError) setValidationError(null);
                  }}
                  placeholder={t(
                    "spreadsheet.validationDependentOptionsPlaceholder",
                  )}
                  rows={3}
                  data-testid="spreadsheet-validation-dep-options"
                  style={{ resize: "vertical" }}
                />
                <input
                  value={valDepDefault}
                  onChange={(event) => {
                    setValDepDefault(event.target.value);
                    if (validationError) setValidationError(null);
                  }}
                  placeholder={t(
                    "spreadsheet.validationDependentDefaultPlaceholder",
                  )}
                  data-testid="spreadsheet-validation-dep-default"
                />
              </span>
            ) : null}
            <label
              className={dialogParts.check}
            >
              <input
                type="checkbox"
                checked={valApplyToSame}
                disabled={
                  !validationOpenedRule ||
                  spreadsheetNativeWriteRoute === "patchback"
                }
                onChange={(event) => setValApplyToSame(event.target.checked)}
                data-testid="spreadsheet-validation-apply-to-same"
              />
              {t("spreadsheet.validationApplyToSame")}
            </label>
            </span>
            {/* Excel-style input message (shown on selection) + error alert
                (title/body/style shown on an invalid entry). All additive — a
                rule left blank here carries no messages. */}
            <span
              id="spreadsheet-validation-panel-input"
              role="tabpanel"
              aria-labelledby="spreadsheet-validation-tab-input"
              data-testid="spreadsheet-validation-panel-input"
              style={{
                display: validationDialogTab === "input" ? "flex" : "none",
                flexDirection: "column",
                gap: 8,
              }}
            >
              <label
                className={dialogParts.check}
              >
                <input
                  type="checkbox"
                  checked={valShowInputMessage}
                  onChange={(event) =>
                    setValShowInputMessage(event.target.checked)
                  }
                  data-testid="spreadsheet-validation-show-input-message"
                />
                {t("spreadsheet.validationShowInputMessage")}
              </label>
              <span
                className={dialogParts.sectionTitle}
              >
                {t("spreadsheet.validationMessagesLabel")}
              </span>
              <input
                value={valInputTitle}
                onChange={(event) => setValInputTitle(event.target.value)}
                disabled={!valShowInputMessage}
                placeholder={t("spreadsheet.validationInputTitle")}
                data-testid="spreadsheet-validation-input-title"
              />
              <input
                value={valInputMessage}
                onChange={(event) => setValInputMessage(event.target.value)}
                disabled={!valShowInputMessage}
                placeholder={t("spreadsheet.validationInputMessage")}
                data-testid="spreadsheet-validation-input-message"
              />
            </span>
            <span
              id="spreadsheet-validation-panel-error"
              role="tabpanel"
              aria-labelledby="spreadsheet-validation-tab-error"
              data-testid="spreadsheet-validation-panel-error"
              style={{
                display: validationDialogTab === "error" ? "flex" : "none",
                flexDirection: "column",
                gap: 8,
              }}
            >
              <label
                className={dialogParts.check}
              >
                <input
                  type="checkbox"
                  checked={valShowErrorMessage}
                  onChange={(event) =>
                    setValShowErrorMessage(event.target.checked)
                  }
                  data-testid="spreadsheet-validation-show-error-message"
                />
                {t("spreadsheet.validationShowErrorMessage")}
              </label>
              <input
                value={valErrorTitle}
                onChange={(event) => setValErrorTitle(event.target.value)}
                disabled={!valShowErrorMessage}
                placeholder={t("spreadsheet.validationErrorTitle")}
                data-testid="spreadsheet-validation-error-title"
              />
              <input
                value={valErrorMessage}
                onChange={(event) => setValErrorMessage(event.target.value)}
                disabled={!valShowErrorMessage}
                placeholder={t("spreadsheet.validationErrorMessage")}
                data-testid="spreadsheet-validation-error-message"
              />
              <label
                className={dialogParts.row}
              >
                {t("spreadsheet.validationErrorStyle")}
                <select
                  value={valErrorStyle}
                  disabled={!valShowErrorMessage}
                  onChange={(event) =>
                    setValErrorStyle(event.target.value as ErrorStyle)
                  }
                  data-testid="spreadsheet-validation-error-style"
                  style={{ flex: 1 }}
                >
                  <option value="stop">
                    {t("spreadsheet.validationErrorStyleStop")}
                  </option>
                  <option value="warning">
                    {t("spreadsheet.validationErrorStyleWarning")}
                  </option>
                  <option value="information">
                    {t("spreadsheet.validationErrorStyleInformation")}
                  </option>
                </select>
              </label>
            </span>
            <span
              id="spreadsheet-validation-panel-ime"
              role="tabpanel"
              aria-labelledby="spreadsheet-validation-tab-ime"
              data-testid="spreadsheet-validation-panel-ime"
              className={dialogParts.hint}
              style={{ display: validationDialogTab === "ime" ? "block" : "none" }}
            >
              {validationImeUnavailableReason}
            </span>
            {validationError ? (
              <span
                data-testid="spreadsheet-validation-error"
                className={dialogParts.error}
              >
                {validationError}
              </span>
            ) : null}
            {(() => {
              // Why OK is unavailable, said where the old button's tooltip
              // was the only place it appeared.
              const applyState = visibleCommandState(
                "XLSCTL.DATA.VALIDATION.MRAI.APPLY",
              );
              const draftBlocker =
                spreadsheetNativeWriteRoute === "patchback" &&
                !validationClearPending &&
                validationDraftRule !== null
                  ? spreadsheetDataValidationPatchbackRuleBlocker(validationDraftRule)
                  : null;
              const reason = applyState.enabled
                ? undefined
                : draftBlocker
                  ? validationPatchbackBlockerMessage(draftBlocker)
                  : visibleCommandReason(applyState.reasonKey);
              return reason ? (
                <span
                  data-testid="spreadsheet-validation-apply-reason"
                  className={dialogParts.hint}
                >
                  {reason}
                </span>
              ) : null;
            })()}
          </span>
        </OfficeDialogShell>
  );
}
