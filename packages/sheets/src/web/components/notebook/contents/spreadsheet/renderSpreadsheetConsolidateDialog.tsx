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
import { type ConsolidateFunction, CONSOLIDATE_FUNCTION_LABEL_KEYS, CONSOLIDATE_FUNCTIONS } from "../../../../lib/spreadsheet/consolidate";
import { type Dispatch, type SetStateAction } from "react";
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import { type SpreadsheetRangePick } from "./useSpreadsheetRangePick";

/** What renderSpreadsheetConsolidateDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetConsolidateDialogInputs {
  rangePick: SpreadsheetRangePick;
  closeConsolidate: () => void;
  consolidateCreateLinks: boolean;
  consolidateDest: string;
  consolidateFn: ConsolidateFunction;
  consolidateLeftCol: boolean;
  consolidateMode: "position" | "category";
  consolidateSources: string;
  consolidateStatus: { kind: "success" | "error"; text: string; } | null;
  consolidateTopRow: boolean;
  consolidateTriggerRef: RefObject<HTMLButtonElement | null>;
  runConsolidate: () => void;
  setConsolidateCreateLinks: Dispatch<SetStateAction<boolean>>;
  setConsolidateDest: Dispatch<SetStateAction<string>>;
  setConsolidateFn: Dispatch<SetStateAction<ConsolidateFunction>>;
  setConsolidateLeftCol: Dispatch<SetStateAction<boolean>>;
  setConsolidateMode: Dispatch<SetStateAction<"position" | "category">>;
  setConsolidateSources: Dispatch<SetStateAction<string>>;
  setConsolidateStatus: Dispatch<SetStateAction<{ kind: "success" | "error"; text: string; } | null>>;
  setConsolidateTopRow: Dispatch<SetStateAction<boolean>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetConsolidateDialog({
  rangePick,
  closeConsolidate,
  consolidateCreateLinks,
  consolidateDest,
  consolidateFn,
  consolidateLeftCol,
  consolidateMode,
  consolidateSources,
  consolidateStatus,
  consolidateTopRow,
  consolidateTriggerRef,
  runConsolidate,
  setConsolidateCreateLinks,
  setConsolidateDest,
  setConsolidateFn,
  setConsolidateLeftCol,
  setConsolidateMode,
  setConsolidateSources,
  setConsolidateStatus,
  setConsolidateTopRow,
  t,
}: RenderSpreadsheetConsolidateDialogInputs) {
  // Collapsed while a range of it is picked on the sheet.
  if (rangePick.picking("spreadsheet-consolidate")) return null;
  return (
        <OfficeDialogShell
          open
          title={t("spreadsheet.consolidateMenu")}
          description={null}
          testIdPrefix="spreadsheet-consolidate"
          defaultActionTestId="spreadsheet-consolidate-run"
          defaultActionLabel={t("spreadsheet.fmtOk")}
          cancelActionLabel={t("common.close")}
          onDefaultAction={runConsolidate}
          onCancel={closeConsolidate}
          returnFocusRef={consolidateTriggerRef}
          className="sm:max-w-[420px] max-h-[calc(100dvh-32px)] overflow-y-auto"
        >
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>
              {t("spreadsheet.consolidateSourcesLabel")}
            </span>
            <span style={{ display: "flex", gap: 4, alignItems: "flex-start" }}>
              <textarea
                value={consolidateSources}
                onChange={(event) => {
                  setConsolidateSources(event.target.value);
                  if (consolidateStatus) setConsolidateStatus(null);
                }}
                rows={4}
                placeholder={t("spreadsheet.consolidateSourcesPlaceholder")}
                data-testid="spreadsheet-consolidate-sources"
                style={{ resize: "vertical", flex: 1, minWidth: 0 }}
              />
              <SpreadsheetCollapseDialogButton
                label={t("spreadsheet.cfDialog.collapse")}
                onClick={() => {
                  // Excel's Reference box: the picked range (on any sheet) is
                  // the next reference, added under the ones already listed.
                  const listed = String(consolidateSources).replace(/\s+$/, "");
                  const head = listed ? `${listed}\n` : "";
                  rangePick.start({
                    id: "spreadsheet-consolidate-sources",
                    dialogLabel: t("spreadsheet.consolidateMenu"),
                    inputLabel: t("spreadsheet.consolidateSourcesLabel"),
                    style: "sheetAbsolute",
                    otherSheets: true,
                    value: "",
                    onChange: (text: string) => {
                      setConsolidateSources(`${head}${text}`);
                      if (consolidateStatus) setConsolidateStatus(null);
                    },
                  });
                }}
                testId="spreadsheet-consolidate-sources-pick"
              />
            </span>
          </label>
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>
              {t("spreadsheet.consolidateFnLabel")}
            </span>
            <select
              value={consolidateFn}
              onChange={(event) => {
                setConsolidateFn(event.target.value as ConsolidateFunction);
                if (consolidateStatus) setConsolidateStatus(null);
              }}
              data-testid="spreadsheet-consolidate-fn"
            >
              {CONSOLIDATE_FUNCTIONS.map((fn) => (
                <option key={fn} value={fn}>
                  {t(CONSOLIDATE_FUNCTION_LABEL_KEYS[fn])}
                </option>
              ))}
            </select>
          </label>
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>
              {t("spreadsheet.consolidateModeLabel")}
            </span>
            <select
              value={consolidateMode}
              onChange={(event) => {
                setConsolidateMode(
                  event.target.value as "position" | "category",
                );
                if (consolidateStatus) setConsolidateStatus(null);
              }}
              data-testid="spreadsheet-consolidate-mode"
            >
              <option value="position">
                {t("spreadsheet.consolidateModePosition")}
              </option>
              <option value="category">
                {t("spreadsheet.consolidateModeCategory")}
              </option>
            </select>
          </label>
          {consolidateMode === "category" ? (
            <span style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label className={dialogParts.check}>
                <input
                  type="checkbox"
                  checked={consolidateTopRow}
                  onChange={(event) => {
                    setConsolidateTopRow(event.target.checked);
                    if (consolidateStatus) setConsolidateStatus(null);
                  }}
                  data-testid="spreadsheet-consolidate-top-row"
                />
                {t("spreadsheet.consolidateTopRow")}
              </label>
              <label className={dialogParts.check}>
                <input
                  type="checkbox"
                  checked={consolidateLeftCol}
                  onChange={(event) => {
                    setConsolidateLeftCol(event.target.checked);
                    if (consolidateStatus) setConsolidateStatus(null);
                  }}
                  data-testid="spreadsheet-consolidate-left-col"
                />
                {t("spreadsheet.consolidateLeftColumn")}
              </label>
            </span>
          ) : null}
          <label
            title={t("spreadsheet.uiCopy.sourcesMustBeInThisWorkbookAndOnSheets")}
            className={dialogParts.check}
          >
            <input
              type="checkbox"
              checked={consolidateCreateLinks}
              onChange={(event) => {
                setConsolidateCreateLinks(event.target.checked);
                if (consolidateStatus) setConsolidateStatus(null);
              }}
              data-testid="spreadsheet-consolidate-create-links"
            />
            {t("spreadsheet.uiCopy.createLinksToSourceData")}
          </label>
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>
              {t("spreadsheet.consolidateDestLabel")}
            </span>
            <span style={{ display: "flex", gap: 4 }}>
              <input
                type="text"
                value={consolidateDest}
                onChange={(event) => {
                  setConsolidateDest(event.target.value);
                  if (consolidateStatus) setConsolidateStatus(null);
                }}
                placeholder={t("spreadsheet.consolidateDestPlaceholder")}
                data-testid="spreadsheet-consolidate-dest"
                style={{ flex: 1, minWidth: 0 }}
              />
              <SpreadsheetCollapseDialogButton
                label={t("spreadsheet.cfDialog.collapse")}
                onClick={() =>
                  rangePick.start({
                    id: "spreadsheet-consolidate-dest",
                    dialogLabel: t("spreadsheet.consolidateMenu"),
                    inputLabel: t("spreadsheet.consolidateDestLabel"),
                    style: "absolute",
                    value: String(consolidateDest),
                    onChange: (text: string) => {
                      setConsolidateDest(text);
                      if (consolidateStatus) setConsolidateStatus(null);
                    },
                  })
                }
                testId="spreadsheet-consolidate-dest-pick"
              />
            </span>
          </label>
          {consolidateStatus ? (
            <span
              role="status"
              data-testid="spreadsheet-consolidate-status"
              data-status={consolidateStatus.kind}
              className={dialogParts.message}
              data-tone={consolidateStatus.kind === "success" ? "success" : "error"}
            >
              {consolidateStatus.text}
            </span>
          ) : null}
        </OfficeDialogShell>
  );
}
