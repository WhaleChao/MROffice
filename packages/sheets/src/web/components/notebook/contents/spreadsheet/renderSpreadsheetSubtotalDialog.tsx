"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { type SubtotalFunction } from "../../../../lib/spreadsheet/subtotal";
import { columnLabel } from "./spreadsheetRangeA1";
import { type GridRect } from "@mrnote/core/web/lib/spreadsheet/grid-nav.ts";
import { type Dispatch, type SetStateAction } from "react";
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

/** What renderSpreadsheetSubtotalDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetSubtotalDialogInputs {
  gridActiveCellFocusTarget: () => HTMLElement | null;
  removeAllSubtotals: () => void;
  runSubtotal: () => void;
  selRect: GridRect;
  setSubtotalFunction: Dispatch<SetStateAction<SubtotalFunction>>;
  setSubtotalGroupCol: Dispatch<SetStateAction<number | null>>;
  setSubtotalOpen: Dispatch<SetStateAction<boolean>>;
  setSubtotalOptions: Dispatch<SetStateAction<{ replace: boolean; pageBreaks: boolean; summaryBelow: boolean; }>>;
  setSubtotalStatus: Dispatch<SetStateAction<{ kind: "success" | "error"; text: string; } | null>>;
  setSubtotalValueCols: Dispatch<SetStateAction<number[]>>;
  subtotalColumnLabel: (column: number) => string;
  subtotalColumns: number[];
  subtotalFunction: SubtotalFunction;
  subtotalGroupCol: number | null;
  subtotalOptions: { replace: boolean; pageBreaks: boolean; summaryBelow: boolean; };
  subtotalStatus: { kind: "success" | "error"; text: string; } | null;
  subtotalValueCols: number[];
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetSubtotalDialog({
  gridActiveCellFocusTarget,
  removeAllSubtotals,
  runSubtotal,
  selRect,
  setSubtotalFunction,
  setSubtotalGroupCol,
  setSubtotalOpen,
  setSubtotalOptions,
  setSubtotalStatus,
  setSubtotalValueCols,
  subtotalColumnLabel,
  subtotalColumns,
  subtotalFunction,
  subtotalGroupCol,
  subtotalOptions,
  subtotalStatus,
  subtotalValueCols,
  t,
}: RenderSpreadsheetSubtotalDialogInputs) {
  return (
        <OfficeDialogShell
          open
          title={t("spreadsheet.subtotalMenu")}
          description={t("spreadsheet.subtotalHint")}
          testIdPrefix="spreadsheet-subtotal"
          defaultActionTestId="spreadsheet-subtotal-run"
          defaultActionLabel={t("spreadsheet.fmtOk")}
          cancelActionLabel={t("spreadsheet.goalSeekCancel")}
          onDefaultAction={runSubtotal}
          onCancel={() => {
            setSubtotalOpen(false);
            setSubtotalStatus(null);
          }}
          returnFocusTarget={gridActiveCellFocusTarget}
          // Remove All sits at the dialog's bottom left, as in Excel, and is
          // the one way to take subtotals out again (XLS-A064).
          footerStart={
            <button
              type="button"
              className={dialogParts.button}
              onClick={removeAllSubtotals}
              data-testid="spreadsheet-subtotal-remove-all"
            >
              <SpreadsheetDialogIcon name="clear" />
              {t("spreadsheet.subtotalRemoveAll")}
            </button>
          }
          className="sm:max-w-[420px] max-h-[calc(100dvh-32px)] overflow-y-auto"
        >
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>
              {t("spreadsheet.subtotalGroupBy")}
            </span>
            <select
              value={subtotalGroupCol ?? selRect.c1}
              onChange={(event) => {
                setSubtotalGroupCol(Number(event.target.value));
                if (subtotalStatus) setSubtotalStatus(null);
              }}
              data-testid="spreadsheet-subtotal-group-col"
            >
              {subtotalColumns.map((c) => (
                <option key={c} value={c}>
                  {subtotalColumnLabel(c)}
                </option>
              ))}
            </select>
          </label>
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>
              {t("spreadsheet.subtotalFunction")}
            </span>
            <select
              value={subtotalFunction}
              onChange={(event) => {
                setSubtotalFunction(event.target.value as SubtotalFunction);
                if (subtotalStatus) setSubtotalStatus(null);
              }}
              data-testid="spreadsheet-subtotal-function"
            >
              {(
                [
                  "sum",
                  "count",
                  "average",
                  "max",
                  "min",
                  "product",
                  "countNumbers",
                  "stdDev",
                  "stdDevp",
                  "var",
                  "varp",
                ] as SubtotalFunction[]
              ).map((fn) => (
                <option key={fn} value={fn}>
                  {t(`spreadsheet.subtotalFn.${fn}`)}
                </option>
              ))}
            </select>
          </label>
          <span className={dialogParts.label}>
            {t("spreadsheet.subtotalValueColumns")}
          </span>
          <span
            className={dialogParts.list}
            style={{ gap: 4, maxHeight: 120, padding: "4px 8px" }}
          >
            {subtotalColumns.map((c) => (
              <label key={c} className={dialogParts.check}>
                <input
                  type="checkbox"
                  checked={subtotalValueCols.includes(c)}
                  onChange={(event) => {
                    setSubtotalValueCols((prev) =>
                      event.target.checked
                        ? [...prev, c].sort((a, b) => a - b)
                        : prev.filter((x) => x !== c),
                    );
                    if (subtotalStatus) setSubtotalStatus(null);
                  }}
                  data-testid={`spreadsheet-subtotal-value-col-${c}`}
                />
                {t("spreadsheet.removeDuplicatesColumn", {
                  column: columnLabel(c),
                })}
              </label>
            ))}
          </span>
          {/* Excel's three options under the value columns (XLS-A064). */}
          {(
            [
              ["replace", "spreadsheet.subtotalReplace"],
              ["pageBreaks", "spreadsheet.subtotalPageBreaks"],
              ["summaryBelow", "spreadsheet.subtotalSummaryBelow"],
            ] as const
          ).map(([option, key]) => (
            <label key={option} className={dialogParts.check}>
              <input
                type="checkbox"
                checked={subtotalOptions[option]}
                onChange={(event) => {
                  const checked = event.currentTarget.checked;
                  setSubtotalOptions((current) => ({ ...current, [option]: checked }));
                  if (subtotalStatus) setSubtotalStatus(null);
                }}
                data-testid={`spreadsheet-subtotal-option-${option}`}
              />
              {t(key)}
            </label>
          ))}
          {subtotalStatus ? (
            <span
              data-testid="spreadsheet-subtotal-status"
              data-status={subtotalStatus.kind}
              className={dialogParts.message}
              data-tone={subtotalStatus.kind === "success" ? "success" : "error"}
            >
              {subtotalStatus.text}
            </span>
          ) : null}
        </OfficeDialogShell>
  );
}
