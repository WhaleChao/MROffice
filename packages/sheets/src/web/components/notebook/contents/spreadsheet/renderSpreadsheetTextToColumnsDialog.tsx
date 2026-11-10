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
import SpreadsheetFixedWidthBreakRuler from "./SpreadsheetFixedWidthBreakRuler";
import { type DelimitedImportDateOrder, type DelimitedImportDecimalSeparator, type DelimitedImportThousandsSeparator, delimitedImportNumberOptionsAreValid, type DelimitedImportNumberOptions } from "../../../../lib/spreadsheet/csv-import";
import { type TextToColumnsQualifier, type TextToColumnsColumnFormat } from "@mrnote/core/web/lib/spreadsheet/text-to-columns.ts";
import { parseCsvImportFixedWidthBreaks } from "./spreadsheetDelimitedImport";
import { type GridRect } from "@mrnote/core/web/lib/spreadsheet/grid-nav.ts";
import { type Dispatch, type SetStateAction } from "react";
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import { type SpreadsheetRangePick } from "./useSpreadsheetRangePick";
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

/** What renderSpreadsheetTextToColumnsDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetTextToColumnsDialogInputs {
  rangePick: SpreadsheetRangePick;
  closeTextToColumns: () => void;
  dataDialogWriteAllowedRef: RefObject<boolean>;
  gridActiveCellFocusTarget: () => HTMLElement | null;
  locale: string;
  normalizedTextToColumnsBreaks: number[];
  rootRef: RefObject<HTMLDivElement | null>;
  runTextToColumns: () => void;
  selRect: GridRect;
  setTextToColumnsBreaks: Dispatch<SetStateAction<string>>;
  setTextToColumnsColumnDateOrder: (index: number, dateOrder: DelimitedImportDateOrder) => void;
  setTextToColumnsColumnFormat: (index: number, format: TextToColumnsColumnFormat) => void;
  setTextToColumnsConsecutive: Dispatch<SetStateAction<boolean>>;
  setTextToColumnsCustom: Dispatch<SetStateAction<string>>;
  setTextToColumnsDelimiterEnabled: (key: "comma" | "custom" | "tab" | "semicolon" | "space", checked: boolean) => void;
  setTextToColumnsDestination: Dispatch<SetStateAction<string>>;
  setTextToColumnsMode: Dispatch<SetStateAction<"delimited" | "fixedWidth">>;
  setTextToColumnsQualifier: Dispatch<SetStateAction<TextToColumnsQualifier>>;
  setTextToColumnsSelectedColumn: Dispatch<SetStateAction<number>>;
  setTextToColumnsStatus: Dispatch<SetStateAction<{ kind: "success" | "error"; text: string; } | null>>;
  setTextToColumnsStep: Dispatch<SetStateAction<1 | 2 | 3>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  textToColumnsBreaks: string;
  textToColumnsColumnDateOrders: DelimitedImportDateOrder[];
  textToColumnsColumnFormats: TextToColumnsColumnFormat[];
  textToColumnsConsecutive: boolean;
  textToColumnsCustom: string;
  textToColumnsDelimiters: { comma: boolean; tab: boolean; semicolon: boolean; space: boolean; custom: boolean; };
  textToColumnsDestination: string;
  textToColumnsMode: "delimited" | "fixedWidth";
  textToColumnsNumberOptions: DelimitedImportNumberOptions;
  textToColumnsPreview: { rows: string[][]; width: number; };
  textToColumnsQualifier: TextToColumnsQualifier;
  textToColumnsReturnToGridRef: RefObject<boolean>;
  textToColumnsSelectedColumn: number;
  textToColumnsSourceLines: string[];
  textToColumnsStatus: { kind: "success" | "error"; text: string; } | null;
  textToColumnsStep: 1 | 2 | 3;
  textToColumnsToggleRef: RefObject<HTMLButtonElement | null>;
  updateTextToColumnsFixedWidthBreaks: (breaks: readonly number[]) => void;
  updateTextToColumnsNumberOptions: (changes: Partial<DelimitedImportNumberOptions>) => void;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetTextToColumnsDialog({
  rangePick,
  closeTextToColumns,
  dataDialogWriteAllowedRef,
  gridActiveCellFocusTarget,
  locale,
  normalizedTextToColumnsBreaks,
  rootRef,
  runTextToColumns,
  selRect,
  setTextToColumnsBreaks,
  setTextToColumnsColumnDateOrder,
  setTextToColumnsColumnFormat,
  setTextToColumnsConsecutive,
  setTextToColumnsCustom,
  setTextToColumnsDelimiterEnabled,
  setTextToColumnsDestination,
  setTextToColumnsMode,
  setTextToColumnsQualifier,
  setTextToColumnsSelectedColumn,
  setTextToColumnsStatus,
  setTextToColumnsStep,
  t,
  textToColumnsBreaks,
  textToColumnsColumnDateOrders,
  textToColumnsColumnFormats,
  textToColumnsConsecutive,
  textToColumnsCustom,
  textToColumnsDelimiters,
  textToColumnsDestination,
  textToColumnsMode,
  textToColumnsNumberOptions,
  textToColumnsPreview,
  textToColumnsQualifier,
  textToColumnsReturnToGridRef,
  textToColumnsSelectedColumn,
  textToColumnsSourceLines,
  textToColumnsStatus,
  textToColumnsStep,
  textToColumnsToggleRef,
  updateTextToColumnsFixedWidthBreaks,
  updateTextToColumnsNumberOptions,
}: RenderSpreadsheetTextToColumnsDialogInputs) {
  // Collapsed while a range of it is picked on the sheet.
  if (rangePick.picking("spreadsheet-text-to-columns")) return null;
  return (
        <OfficeDialogShell
          open
          // XLS-A074: Excel's three-step Convert Text to Columns Wizard.
          title={t("spreadsheet.uiCopy.textToColumnsWizardTitle", { step: textToColumnsStep })}
          description={
            textToColumnsStep === 1
              ? t("spreadsheet.uiCopy.textToColumnsStep1Intro")
              : textToColumnsStep === 2
                ? textToColumnsMode === "fixedWidth"
                  ? t("spreadsheet.uiCopy.textToColumnsStep2FixedIntro")
                  : t("spreadsheet.uiCopy.textToColumnsStep2DelimitedIntro")
                : t("spreadsheet.uiCopy.textToColumnsStep3Intro")
          }
          footerStart={
            <span className="inline-flex gap-2">
              <button
                type="button"
                className={dialogParts.button}
                disabled={textToColumnsStep === 1}
                onClick={() => setTextToColumnsStep((step) => (step > 1 ? ((step - 1) as 1 | 2) : step))}
                data-testid="spreadsheet-text-to-columns-back"
              >
                <SpreadsheetDialogIcon name="back" />
                {t("spreadsheet.uiCopy.wizardBack")}
              </button>
              <button
                type="button"
                className={dialogParts.button}
                disabled={textToColumnsStep === 3}
                onClick={() => setTextToColumnsStep((step) => (step < 3 ? ((step + 1) as 2 | 3) : step))}
                data-testid="spreadsheet-text-to-columns-next"
              >
                {t("spreadsheet.uiCopy.wizardNext")}
                <SpreadsheetDialogIcon name="next" />
              </button>
            </span>
          }
          testIdPrefix="spreadsheet-text-to-columns"
          defaultActionTestId="spreadsheet-text-to-columns-run"
          defaultActionLabel={t("spreadsheet.uiCopy.finish")}
          cancelActionLabel={t("spreadsheet.goalSeekCancel")}
          onDefaultAction={runTextToColumns}
          onCancel={closeTextToColumns}
          defaultActionDisabled={!dataDialogWriteAllowedRef.current}
          returnFocusTarget={() => {
            if (textToColumnsReturnToGridRef.current) {
              textToColumnsReturnToGridRef.current = false;
              return gridActiveCellFocusTarget();
            }
            return rootRef.current?.querySelector<HTMLElement>(
              '[data-testid="spreadsheet-ribbon-data-text-to-columns"]',
            ) ?? textToColumnsToggleRef.current;
          }}
          className="max-h-[calc(100dvh-32px)] overflow-y-auto"
        >
          {textToColumnsStep === 1 ? (
            <>
              <fieldset
                className="m-0 grid gap-1.5 border-0 p-0"
                data-testid="spreadsheet-text-to-columns-mode"
              >
                <legend className={dialogParts.sectionTitle}>
                  {t("spreadsheet.uiCopy.textToColumnsOriginalDataType")}
                </legend>
                {(
                  [
                    ["delimited", "spreadsheet.textToColumnsModeDelimited", "spreadsheet.uiCopy.textToColumnsDelimitedHelp"],
                    ["fixedWidth", "spreadsheet.textToColumnsModeFixedWidth", "spreadsheet.uiCopy.textToColumnsFixedWidthHelp"],
                  ] as const
                ).map(([value, label, help]) => (
                  <label key={value} className={dialogParts.check}>
                    <input
                      type="radio"
                      name="spreadsheet-text-to-columns-mode"
                      checked={textToColumnsMode === value}
                      onChange={() => {
                        setTextToColumnsMode(value);
                        if (textToColumnsStatus) setTextToColumnsStatus(null);
                      }}
                      data-testid={`spreadsheet-text-to-columns-mode-${value}`}
                    />
                    <SpreadsheetDialogIcon name={value} />
                    <span>
                      <strong className="font-medium">{t(label)}</strong>
                      {" - "}
                      {t(help)}
                    </span>
                  </label>
                ))}
              </fieldset>
              <span className={dialogParts.label}>
                {t("spreadsheet.uiCopy.textToColumnsPreviewOfSelected")}
              </span>
              <pre
                data-testid="spreadsheet-text-to-columns-source-preview"
                className={`${dialogParts.well} m-0 max-h-32 overflow-auto px-2 py-1 font-mono text-[11px]`}
              >
                {textToColumnsSourceLines
                  .slice(0, 8)
                  .map((line, index) => `${selRect.r1 + index + 1} ${line}`)
                  .join("\n")}
              </pre>
            </>
          ) : null}
          {textToColumnsStep !== 2 ? null : textToColumnsMode === "delimited" ? (
            <fieldset
              data-testid="spreadsheet-text-to-columns-delimiter"
              style={{
                display: "flex",
                flexDirection: "column",
                gap: 6,
                margin: 0,
                padding: 0,
                border: 0,
              }}
            >
              <legend
                style={{
                  fontSize: 11,
                  color: "var(--muted, #6b7280)",
                }}
              >
                {t(
                  "spreadsheet.textToColumnsDelimiterLabel",
                )}
              </legend>
              <span
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  alignItems: "center",
                  gap: "4px 12px",
                }}
              >
                {(
                  [
                    ["tab", "spreadsheet.textToColumnsTab"],
                    [
                      "semicolon",
                      "spreadsheet.textToColumnsSemicolon",
                    ],
                    [
                      "comma",
                      "spreadsheet.textToColumnsComma",
                    ],
                    [
                      "space",
                      "spreadsheet.textToColumnsSpace",
                    ],
                    [
                      "custom",
                      "spreadsheet.textToColumnsCustom",
                    ],
                  ] as const
                ).map(([key, label]) => (
                  <label
                    key={key}
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 4,
                      fontSize: 11,
                      whiteSpace: "nowrap",
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={
                        textToColumnsDelimiters[key]
                      }
                      onChange={(event) =>
                        setTextToColumnsDelimiterEnabled(
                          key,
                          event.target.checked,
                        )
                      }
                      data-testid={`spreadsheet-text-to-columns-delimiter-${key}`}
                    />
                    <span>{t(label)}</span>
                  </label>
                ))}
              </span>
              {textToColumnsDelimiters.custom ? (
                <label
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: 2,
                  }}
                >
                  <span
                    style={{
                      fontSize: 11,
                      color: "var(--muted, #6b7280)",
                    }}
                  >
                    {t(
                      "spreadsheet.textToColumnsCustom",
                    )}
                  </span>
                  <input
                    value={textToColumnsCustom}
                    maxLength={1}
                    onChange={(event) => {
                      setTextToColumnsCustom(
                        event.target.value,
                      );
                      setTextToColumnsStatus(null);
                    }}
                    placeholder={t(
                      "spreadsheet.textToColumnsCustomPlaceholder",
                    )}
                    data-testid="spreadsheet-text-to-columns-custom"
                    style={{
                      fontSize: 12,
                      padding: "4px 6px",
                      border:
                        "1px solid var(--border, #e5e7eb)",
                      borderRadius: 5,
                    }}
                  />
                </label>
              ) : null}
              <span
                style={{
                  display: "grid",
                  gridTemplateColumns:
                    "minmax(0, 1fr) minmax(0, 1fr)",
                  gap: 8,
                  alignItems: "end",
                }}
              >
                <label
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: 2,
                  }}
                >
                  <span
                    style={{
                      fontSize: 11,
                      color: "var(--muted, #6b7280)",
                    }}
                  >
                    {t(
                      "spreadsheet.textToColumnsQualifier",
                    )}
                  </span>
                  <select
                    value={textToColumnsQualifier}
                    onChange={(event) => {
                      setTextToColumnsQualifier(
                        event.target
                          .value as TextToColumnsQualifier,
                      );
                      setTextToColumnsStatus(null);
                    }}
                    data-testid="spreadsheet-text-to-columns-qualifier"
                    style={{
                      fontSize: 12,
                      padding: "4px 6px",
                      border:
                        "1px solid var(--border, #e5e7eb)",
                      borderRadius: 5,
                    }}
                  >
                    <option value="double">
                      {t(
                        "spreadsheet.textToColumnsQualifierDouble",
                      )}
                    </option>
                    <option value="single">
                      {t(
                        "spreadsheet.textToColumnsQualifierSingle",
                      )}
                    </option>
                    <option value="none">
                      {t(
                        "spreadsheet.textToColumnsQualifierNone",
                      )}
                    </option>
                  </select>
                </label>
                <label
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 4,
                    minHeight: 28,
                    fontSize: 11,
                  }}
                >
                  <input
                    type="checkbox"
                    checked={textToColumnsConsecutive}
                    onChange={(event) => {
                      setTextToColumnsConsecutive(
                        event.target.checked,
                      );
                      setTextToColumnsStatus(null);
                    }}
                    data-testid="spreadsheet-text-to-columns-consecutive"
                  />
                  <span>
                    {t(
                      "spreadsheet.textToColumnsConsecutive",
                    )}
                  </span>
                </label>
              </span>
            </fieldset>
          ) : (
            <div className="grid gap-1.5">
              <label className={dialogParts.field}>
                <span className={dialogParts.label}>
                  {t(
                    "spreadsheet.textToColumnsBreaksLabel",
                  )}
                </span>
                <input
                  value={textToColumnsBreaks}
                  onChange={(event) => {
                    setTextToColumnsBreaks(
                      event.target.value,
                    );
                    if (textToColumnsStatus)
                      setTextToColumnsStatus(null);
                  }}
                  placeholder={t(
                    "spreadsheet.textToColumnsBreaksPlaceholder",
                  )}
                  aria-invalid={
                    parseCsvImportFixedWidthBreaks(
                      textToColumnsBreaks,
                    ) === null
                  }
                  aria-describedby="spreadsheet-text-to-columns-fixed-width-help"
                  data-testid="spreadsheet-text-to-columns-breaks"
                />
              </label>
              <span
                id="spreadsheet-text-to-columns-fixed-width-help"
                className={dialogParts.hint}
              >
                {t("spreadsheet.uiCopy.clickToAddDoubleClickOrDeleteToRemove")}
              </span>
              <SpreadsheetFixedWidthBreakRuler
                lines={textToColumnsSourceLines}
                breaks={normalizedTextToColumnsBreaks}
                onChange={
                  updateTextToColumnsFixedWidthBreaks
                }
                translate={t}
                testIdPrefix="spreadsheet-text-to-columns-fixed-width"
                helpId="spreadsheet-text-to-columns-fixed-width-help"
              />
            </div>
          )}
          {textToColumnsStep === 3 ? (
            <>
          <fieldset
            className="m-0 grid gap-1 border-0 p-0"
            data-testid="spreadsheet-text-to-columns-column-format"
          >
            <legend className={dialogParts.sectionTitle}>
              {t("spreadsheet.uiCopy.textToColumnsColumnDataFormat")}
            </legend>
            {(
              [
                ["general", "spreadsheet.textToColumnsFormatGeneral"],
                ["text", "spreadsheet.textToColumnsFormatText"],
                ["date", "spreadsheet.fmtCatDate"],
                ["skip", "spreadsheet.textToColumnsFormatSkip"],
              ] as const
            ).map(([value, label]) => (
              <span key={value} className="inline-flex items-center gap-2">
                <label className={dialogParts.check}>
                  <input
                    type="radio"
                    name="spreadsheet-text-to-columns-column-format"
                    checked={(textToColumnsColumnFormats[textToColumnsSelectedColumn] ?? "general") === value}
                    onChange={() => setTextToColumnsColumnFormat(textToColumnsSelectedColumn, value)}
                    data-testid={`spreadsheet-text-to-columns-format-${value}`}
                  />
                  <span>{t(label)}</span>
                </label>
                {value === "date" ? (
                  <select
                    value={textToColumnsColumnDateOrders[textToColumnsSelectedColumn] ?? "locale"}
                    disabled={textToColumnsColumnFormats[textToColumnsSelectedColumn] !== "date"}
                    onChange={(event) =>
                      setTextToColumnsColumnDateOrder(
                        textToColumnsSelectedColumn,
                        event.target.value as DelimitedImportDateOrder,
                      )
                    }
                    aria-label={t("spreadsheet.dynamicCopy.columnDateOrder", { index: textToColumnsSelectedColumn + 1 })}
                    data-testid="spreadsheet-text-to-columns-date-order"
                  >
                    <option value="locale">{t("spreadsheet.uiCopy.locale")}</option>
                    <option value="mdy">MDY</option>
                    <option value="dmy">DMY</option>
                    <option value="ymd">YMD</option>
                  </select>
                ) : null}
              </span>
            ))}
          </fieldset>
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>
              {t(
                "spreadsheet.textToColumnsDestination",
              )}
            </span>
            <span style={{ display: "flex", gap: 4 }}>
              <input
                value={textToColumnsDestination}
                onChange={(event) => {
                  setTextToColumnsDestination(
                    event.target.value,
                  );
                  setTextToColumnsStatus(null);
                }}
                placeholder={t(
                  "spreadsheet.textToColumnsDestinationPlaceholder",
                )}
                spellCheck={false}
                data-testid="spreadsheet-text-to-columns-destination"
                style={{ flex: 1, minWidth: 0 }}
              />
              <SpreadsheetCollapseDialogButton
                label={t("spreadsheet.cfDialog.collapse")}
                onClick={() =>
                  rangePick.start({
                    id: "spreadsheet-text-to-columns-destination",
                    dialogLabel: t("spreadsheet.uiCopy.textToColumnsWizardTitle", { step: textToColumnsStep }),
                    inputLabel: t("spreadsheet.textToColumnsDestination"),
                    style: "absolute",
                    value: String(textToColumnsDestination),
                    onChange: (text: string) => {
                      setTextToColumnsDestination(
                        text,
                      );
                      setTextToColumnsStatus(null);
                    },
                  })
                }
                testId="spreadsheet-text-to-columns-destination-pick"
              />
            </span>
          </label>
          <details
            data-testid="spreadsheet-text-to-columns-number-options"
          >
            {/* A disclosure within the page, not a box inside the card. */}
            <summary className={`${dialogParts.sectionTitle} cursor-pointer py-1`}>
              {t("spreadsheet.uiCopy.advancedNumberParsing")}
            </summary>
            <div className="grid grid-cols-2 gap-2 pt-1">
              <label className={dialogParts.field}>
                <span className={dialogParts.label}>
                  {t("spreadsheet.uiCopy.decimalSeparator")}
                </span>
                <select
                  value={
                    textToColumnsNumberOptions.decimalSeparator ??
                    "locale"
                  }
                  onChange={(event) =>
                    updateTextToColumnsNumberOptions({
                      decimalSeparator: event.target
                        .value as DelimitedImportDecimalSeparator,
                    })
                  }
                  data-testid="spreadsheet-text-to-columns-decimal-separator"
                >
                  <option value="locale">
                    {t("spreadsheet.uiCopy.autoLocale")}
                  </option>
                  <option value=".">.</option>
                  <option value=",">,</option>
                </select>
              </label>
              <label className={dialogParts.field}>
                <span className={dialogParts.label}>
                  {t("spreadsheet.uiCopy.thousandsSeparator")}
                </span>
                <select
                  value={
                    textToColumnsNumberOptions.thousandsSeparator ??
                    "locale"
                  }
                  onChange={(event) =>
                    updateTextToColumnsNumberOptions({
                      thousandsSeparator: event.target
                        .value as DelimitedImportThousandsSeparator,
                    })
                  }
                  data-testid="spreadsheet-text-to-columns-thousands-separator"
                >
                  <option value="locale">
                    {t("spreadsheet.uiCopy.autoLocale")}
                  </option>
                  <option value=",">,</option>
                  <option value=".">.</option>
                  <option value="space">
                    {t("spreadsheet.textToColumnsSpace")}
                  </option>
                  <option value="none">
                    {t("spreadsheet.textArt_plain")}
                  </option>
                </select>
              </label>
              <label className={`${dialogParts.check} col-span-2`}>
                <input
                  type="checkbox"
                  checked={Boolean(
                    textToColumnsNumberOptions.trailingMinus,
                  )}
                  onChange={(event) =>
                    updateTextToColumnsNumberOptions({
                      trailingMinus: event.target.checked,
                    })
                  }
                  data-testid="spreadsheet-text-to-columns-trailing-minus"
                />
                <span>
                  {t("spreadsheet.uiCopy.trailingMinusForNegativeNumbers")}
                </span>
              </label>
              <span
                className={`${
                  delimitedImportNumberOptionsAreValid(
                    locale,
                    textToColumnsNumberOptions,
                  )
                    ? dialogParts.hint
                    : dialogParts.error
                } col-span-2`}
              >
                {delimitedImportNumberOptionsAreValid(
                  locale,
                  textToColumnsNumberOptions,
                )
                  ? t("spreadsheet.uiCopy.appliesOnlyToGeneralAndDateColumnsTextRemains")
                  : t("spreadsheet.uiCopy.decimalAndThousandsSeparatorsMustDiffer")}
              </span>
            </div>
          </details>
            </>
          ) : null}
          {textToColumnsStep === 1 ? null : (
          <span
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 4,
            }}
          >
            <span
              style={{
                fontSize: 11,
                color: "var(--muted, #6b7280)",
              }}
            >
              {t("spreadsheet.uiCopy.textToColumnsDataPreview")}
            </span>
            <span
              data-testid="spreadsheet-text-to-columns-preview"
              style={{
                display: "block",
                maxWidth: "100%",
                overflowX: "auto",
                border:
                  "1px solid var(--border, #e5e7eb)",
                borderRadius: 5,
              }}
            >
              <table
                style={{
                  minWidth: "100%",
                  borderCollapse: "collapse",
                  tableLayout: "auto",
                }}
              >
                <thead>
                  <tr>
                    {Array.from(
                      {
                        length:
                          textToColumnsPreview.width,
                      },
                      (_, index) => (
                        <th
                          key={index}
                          scope="col"
                          style={{
                            minWidth: 92,
                            padding: 4,
                            textAlign: "left",
                            background:
                              "var(--surface-muted, #f8fafc)",
                            borderRight:
                              index + 1 <
                              textToColumnsPreview.width
                                ? "1px solid var(--border, #e5e7eb)"
                                : undefined,
                            borderBottom:
                              "1px solid var(--border, #e5e7eb)",
                          }}
                        >
                          {textToColumnsStep === 3 ? (
                            <button
                              type="button"
                              aria-pressed={textToColumnsSelectedColumn === index}
                              onClick={() => setTextToColumnsSelectedColumn(index)}
                              data-testid={`spreadsheet-text-to-columns-column-${index}`}
                              style={{
                                font: "inherit",
                                fontSize: 11,
                                minHeight: 0,
                                padding: "1px 4px",
                                border: 0,
                                borderRadius: 3,
                                background:
                                  textToColumnsSelectedColumn === index
                                    ? "var(--console-accent-strong, #087c73)"
                                    : "transparent",
                                color:
                                  textToColumnsSelectedColumn === index
                                    ? "#fff"
                                    : "inherit",
                                cursor: "pointer",
                              }}
                            >
                              {(() => {
                                const format = textToColumnsColumnFormats[index] ?? "general";
                                if (format === "date") {
                                  const order = textToColumnsColumnDateOrders[index] ?? "locale";
                                  return order === "locale" ? t("spreadsheet.fmtCatDate") : order.toUpperCase();
                                }
                                return t(
                                  format === "text"
                                    ? "spreadsheet.textToColumnsFormatText"
                                    : format === "skip"
                                      ? "spreadsheet.uiCopy.textToColumnsSkipColumn"
                                      : "spreadsheet.textToColumnsFormatGeneral",
                                );
                              })()}
                            </button>
                          ) : (
                            <span style={{ fontSize: 10, color: "var(--muted, #6b7280)" }}>
                              {t("spreadsheet.textToColumnsPreviewColumn", { column: index + 1 })}
                            </span>
                          )}
                        </th>
                      ),
                    )}
                  </tr>
                </thead>
                <tbody>
                  {textToColumnsPreview.rows.map(
                    (row, rowIndex) => (
                      <tr key={rowIndex}>
                        {Array.from(
                          {
                            length:
                              textToColumnsPreview.width,
                          },
                          (_, columnIndex) => (
                            <td
                              key={columnIndex}
                              title={
                                row[columnIndex] ?? ""
                              }
                              style={{
                                maxWidth: 160,
                                height: 24,
                                padding: "4px 6px",
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                // Each field shows its spaces as the cell
                                // will hold them: leading, trailing and runs
                                // of spaces inside quotes (live audit S-109).
                                whiteSpace: "pre",
                                fontSize: 11,
                                borderRight:
                                  columnIndex + 1 <
                                  textToColumnsPreview.width
                                    ? "1px solid var(--border, #e5e7eb)"
                                    : undefined,
                                borderBottom:
                                  rowIndex + 1 <
                                  textToColumnsPreview.rows
                                    .length
                                    ? "1px solid var(--border, #e5e7eb)"
                                    : undefined,
                              }}
                            >
                              {row[columnIndex] ?? ""}
                            </td>
                          ),
                        )}
                      </tr>
                    ),
                  )}
                </tbody>
              </table>
            </span>
          </span>
          )}
          {textToColumnsStatus ? (
            <span
              data-testid="spreadsheet-text-to-columns-status"
              data-status={textToColumnsStatus.kind}
              style={{
                fontSize: 11,
                color:
                  textToColumnsStatus.kind === "success"
                    ? "var(--success, #16a34a)"
                    : "var(--danger, #dc2626)",
              }}
            >
              {textToColumnsStatus.text}
            </span>
          ) : null}
        </OfficeDialogShell>
  );
}
