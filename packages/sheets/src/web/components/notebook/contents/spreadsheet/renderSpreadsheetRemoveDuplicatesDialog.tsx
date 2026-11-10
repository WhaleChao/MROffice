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
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

/** What renderSpreadsheetRemoveDuplicatesDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetRemoveDuplicatesDialogInputs {
  closeRemoveDuplicates: () => void;
  dataDialogWriteAllowedRef: RefObject<boolean>;
  gridActiveCellFocusTarget: () => HTMLElement | null;
  removeDuplicatesColumnLabels: string[];
  removeDuplicatesColumns: number[];
  removeDuplicatesHasHeader: boolean;
  removeDuplicatesKeyCols: Set<number>;
  removeDuplicatesReturnFocusRef: RefObject<HTMLElement | null>;
  removeDuplicatesReturnToGridRef: RefObject<boolean>;
  removeDuplicatesStatus: { kind: "success" | "error"; text: string; } | null;
  rootRef: RefObject<HTMLDivElement | null>;
  runRemoveDuplicates: () => void;
  setRemoveDuplicatesHasHeader: Dispatch<SetStateAction<boolean>>;
  setRemoveDuplicatesKeyCols: Dispatch<SetStateAction<Set<number>>>;
  setRemoveDuplicatesStatus: Dispatch<SetStateAction<{ kind: "success" | "error"; text: string; } | null>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetRemoveDuplicatesDialog({
  closeRemoveDuplicates,
  dataDialogWriteAllowedRef,
  gridActiveCellFocusTarget,
  removeDuplicatesColumnLabels,
  removeDuplicatesColumns,
  removeDuplicatesHasHeader,
  removeDuplicatesKeyCols,
  removeDuplicatesReturnFocusRef,
  removeDuplicatesReturnToGridRef,
  removeDuplicatesStatus,
  rootRef,
  runRemoveDuplicates,
  setRemoveDuplicatesHasHeader,
  setRemoveDuplicatesKeyCols,
  setRemoveDuplicatesStatus,
  t,
}: RenderSpreadsheetRemoveDuplicatesDialogInputs) {
  return (
        <OfficeDialogShell
          open
          title={t("spreadsheet.removeDuplicatesMenu")}
          description={t("spreadsheet.removeDuplicatesColumnsLabel")}
          testIdPrefix="spreadsheet-remove-duplicates"
          defaultActionTestId="spreadsheet-remove-duplicates-run"
          defaultActionLabel={t("spreadsheet.fmtOk")}
          cancelActionLabel={t("spreadsheet.goalSeekCancel")}
          onDefaultAction={runRemoveDuplicates}
          onCancel={closeRemoveDuplicates}
          defaultActionDisabled={!dataDialogWriteAllowedRef.current || removeDuplicatesKeyCols.size === 0}
          returnFocusTarget={() =>
            removeDuplicatesReturnToGridRef.current
              ? gridActiveCellFocusTarget()
              : (rootRef.current?.querySelector<HTMLElement>(
                  '[data-testid="spreadsheet-ribbon-data-remove-duplicates"]',
                ) ?? removeDuplicatesReturnFocusRef.current)
          }
          className="max-h-[calc(100dvh-32px)] overflow-y-auto"
        >
          <label
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              fontSize: 12,
            }}
          >
            <input
              type="checkbox"
              checked={removeDuplicatesHasHeader}
              onChange={(event) => {
                setRemoveDuplicatesHasHeader(
                  event.target.checked,
                );
                if (removeDuplicatesStatus)
                  setRemoveDuplicatesStatus(null);
              }}
              data-testid="spreadsheet-remove-duplicates-header"
            />
            {t("spreadsheet.removeDuplicatesHasHeader")}
          </label>

          <span
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            <button
              type="button"
              onClick={() => {
                setRemoveDuplicatesKeyCols(
                  new Set(removeDuplicatesColumns),
                );
                setRemoveDuplicatesStatus(null);
              }}
              disabled={removeDuplicatesColumns.every(
                (col) =>
                  removeDuplicatesKeyCols.has(col),
              )}
              data-testid="spreadsheet-remove-duplicates-select-all"
              className={dialogParts.compactButton}
            >
              <SpreadsheetDialogIcon name="selectAll" />
              {t("toolbar.selectAll")}
            </button>
            <button
              type="button"
              onClick={() => {
                setRemoveDuplicatesKeyCols(new Set());
                setRemoveDuplicatesStatus(null);
              }}
              disabled={removeDuplicatesKeyCols.size === 0}
              data-testid="spreadsheet-remove-duplicates-unselect-all"
              className={dialogParts.compactButton}
            >
              <SpreadsheetDialogIcon name="deselectAll" />
              {t("library.clearSelection")}
            </button>
          </span>
          {removeDuplicatesColumns.map((col, index) => (
            <label
              key={col}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                fontSize: 12,
              }}
            >
              <input
                type="checkbox"
                checked={removeDuplicatesKeyCols.has(col)}
                onChange={(event) => {
                  setRemoveDuplicatesKeyCols((prev) => {
                    const next = new Set(prev);
                    if (event.target.checked)
                      next.add(col);
                    else next.delete(col);
                    return next;
                  });
                  if (removeDuplicatesStatus) {
                    setRemoveDuplicatesStatus(null);
                  }
                }}
                data-testid={`spreadsheet-remove-duplicates-col-${col}`}
              />
              {removeDuplicatesColumnLabels[index]}
            </label>
          ))}
          {removeDuplicatesStatus ? (
            <span
              data-testid="spreadsheet-remove-duplicates-status"
              data-status={removeDuplicatesStatus.kind}
              style={{
                fontSize: 11,
                color:
                  removeDuplicatesStatus.kind ===
                  "success"
                    ? "var(--success, #16a34a)"
                    : "var(--danger, #dc2626)",
              }}
            >
              {removeDuplicatesStatus.text}
            </span>
          ) : null}
        </OfficeDialogShell>
  );
}
