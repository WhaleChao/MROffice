"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { nativeQueryRefreshable } from "../../../../lib/spreadsheet/native-query-refresh";
import type { SpreadsheetWorkbookQuery } from "../../../../lib/spreadsheet/workbook-import-client";

type Translate = (key: string, values?: Record<string, string | number>) => string;

function lastRefreshedText(query: SpreadsheetWorkbookQuery, locale: string | undefined): string | null {
  if (!query.lastRefreshed) return null;
  const time = Date.parse(query.lastRefreshed.replace(/(\.\d{3})\d*Z$/, "$1Z"));
  if (!Number.isFinite(time)) return null;
  return new Date(time).toLocaleString(locale, { dateStyle: "medium", timeStyle: "short" });
}

/**
 * The workbook's own Power Query queries in Queries & Connections, as Excel's
 * pane lists them: name, where each loads, when it last refreshed, and its
 * source. A query over this workbook's tables refreshes here; one reading an
 * external source stays as Excel saved it (its M text can still be read).
 */
export function SpreadsheetNativeQueriesList(props: {
  translate: Translate;
  locale?: string;
  queries: readonly SpreadsheetWorkbookQuery[];
  /** Why refresh is unavailable on this page (read-only, protected), or null. */
  refreshBlockedReason: string | null;
  busyName: string | null;
  onRefresh: (name: string) => void;
  onViewFormula: (name: string) => void;
}) {
  const { translate: t, queries } = props;
  if (!queries.length) return null;
  return (
    <section className="grid gap-1" aria-label={t("spreadsheet.nativeQuery.queries")}>
      <h3 className={officeDialogParts.sectionTitle}>{t("spreadsheet.nativeQuery.queries")}</h3>
      <ul className={officeDialogParts.list} data-testid="spreadsheet-native-queries">
        {queries.map((query) => {
          const refreshed = lastRefreshedText(query, props.locale);
          const blocked = nativeQueryRefreshable(query);
          const reason =
            blocked === "external"
              ? t("spreadsheet.nativeQuery.externalReason", { functions: query.externalFunctions.join(", ") })
              : blocked === "not_loaded"
                ? t("spreadsheet.nativeQuery.notLoaded", { name: query.name })
                : props.refreshBlockedReason;
          return (
            <li
              key={query.name}
              className={`${officeDialogParts.listItem} flex items-start gap-2`}
              data-testid={`spreadsheet-native-query-${query.name}`}
            >
              <span className="min-w-0 flex-1">
                <span className="block truncate font-medium">{query.name}</span>
                <span className={`${officeDialogParts.hint} block truncate`}>
                  {query.load === "table" && query.tableName
                    ? t("spreadsheet.nativeQuery.loadTable", {
                        sheet: query.sheetName ?? "",
                        table: query.tableName,
                        rows: query.rowCount ?? 0,
                      })
                    : query.load === "data_model"
                      ? t("spreadsheet.nativeQuery.loadDataModel")
                      : t("spreadsheet.nativeQuery.loadConnectionOnly")}
                </span>
                <span className={`${officeDialogParts.hint} block truncate`}>
                  {query.sourceKind === "external"
                    ? t("spreadsheet.nativeQuery.sourceExternal", { functions: query.externalFunctions.join(", ") || "—" })
                    : t("spreadsheet.nativeQuery.sourceLocal")}
                </span>
                {refreshed ? (
                  <span className={`${officeDialogParts.hint} block truncate`}>
                    {t("spreadsheet.nativeQuery.lastRefreshed", { time: refreshed })}
                  </span>
                ) : null}
              </span>
              <button
                type="button"
                className={officeDialogParts.compactButton}
                data-testid={`spreadsheet-native-query-formula-${query.name}`}
                onClick={() => props.onViewFormula(query.name)}
              >
                {t("spreadsheet.nativeQuery.viewFormula")}
              </button>
              <button
                type="button"
                className={officeDialogParts.compactButton}
                disabled={Boolean(reason) || props.busyName !== null}
                title={reason ?? undefined}
                aria-label={reason ? `${t("spreadsheet.nativeQuery.refresh")}: ${reason}` : undefined}
                data-testid={`spreadsheet-native-query-refresh-${query.name}`}
                onClick={() => props.onRefresh(query.name)}
              >
                {props.busyName === query.name ? t("spreadsheet.nativeQuery.refreshing") : t("spreadsheet.nativeQuery.refresh")}
              </button>
            </li>
          );
        })}
      </ul>
    </section>
  );
}

/** Excel's Advanced Editor, read-only: the query's M text as the file has it. */
export function SpreadsheetNativeQueryFormulaDialog(props: {
  translate: Translate;
  query: SpreadsheetWorkbookQuery;
  onClose: () => void;
}) {
  const { translate: t, query } = props;
  const formula = query.formula.replace(/\r\n?/g, "\n");
  return (
    <OfficeDialogShell
      open
      title={t("spreadsheet.nativeQuery.formulaTitle", { name: query.name })}
      description={
        query.sourceKind === "external"
          ? t("spreadsheet.nativeQuery.externalReason", { functions: query.externalFunctions.join(", ") })
          : t("spreadsheet.nativeQuery.formulaDescription")
      }
      defaultActionLabel={t("spreadsheet.nativeQuery.close")}
      cancelActionLabel={t("spreadsheet.nativeQuery.close")}
      showCancelAction={false}
      onDefaultAction={props.onClose}
      onCancel={props.onClose}
      testIdPrefix="spreadsheet-native-query-formula"
    >
      <pre
        className={`${officeDialogParts.well} ${officeDialogParts.mono} max-h-[420px] overflow-auto whitespace-pre p-3 text-xs`}
        data-testid="spreadsheet-native-query-formula-text"
        tabIndex={0}
      >
        {`${query.name.match(/^[A-Za-z_][A-Za-z0-9_]*$/) ? query.name : `#"${query.name.replace(/"/g, '""')}"`} = ${formula}`}
      </pre>
    </OfficeDialogShell>
  );
}
