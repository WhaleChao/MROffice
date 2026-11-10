"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Fragment } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { type Dispatch, type SetStateAction } from "react";

/** What renderSpreadsheetAccessibilityDialog reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetAccessibilityDialogInputs {
  reviewDialogMode: "accessibility" | "performance" | "statistics" | null;
  setReviewDialogMode: Dispatch<SetStateAction<"accessibility" | "performance" | "statistics" | null>>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  workbookReviewStats: { accessibilityIssues: string[]; cellsWithData: number; charts: number; comments: number; conditionalFormats: number; formulas: number; images: number; mergedRanges: number; objects: number; pivots: number; sheets: number; tables: number; } | null;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetAccessibilityDialog({
  reviewDialogMode,
  setReviewDialogMode,
  t,
  workbookReviewStats,
}: RenderSpreadsheetAccessibilityDialogInputs) {
  return (
      <OfficeDialogShell
        open={reviewDialogMode !== null}
        title={
          reviewDialogMode === "accessibility"
            ? t("spreadsheet.review.accessibilityTitle")
            : reviewDialogMode === "performance"
              ? t("spreadsheet.review.performanceTitle")
              : t("spreadsheet.review.statisticsTitle")
        }
        description={
          reviewDialogMode === "accessibility"
            ? t("spreadsheet.review.accessibilityDescription")
            : reviewDialogMode === "performance"
              ? t("spreadsheet.review.performanceDescription")
              : t("spreadsheet.review.statisticsDescription")
        }
        defaultActionLabel={t("common.close")}
        cancelActionLabel={t("common.cancel")}
        // An informational dialog (Excel's Workbook Statistics has one Close);
        // a Cancel beside it did the same thing.
        showCancelAction={false}
        onDefaultAction={() => setReviewDialogMode(null)}
        onCancel={() => setReviewDialogMode(null)}
        testIdPrefix="spreadsheet-review"
      >
        {reviewDialogMode === "accessibility" && workbookReviewStats ? (
          workbookReviewStats.accessibilityIssues.length === 0 ? (
            <p data-testid="spreadsheet-review-accessibility-clean">
              {t("spreadsheet.review.accessibilityClean")}
            </p>
          ) : (
            <ul
              data-testid="spreadsheet-review-accessibility-issues"
              style={{ margin: 0, paddingInlineStart: 20, fontSize: 12 }}
            >
              {workbookReviewStats.accessibilityIssues.map((issue) => (
                <li key={issue}>{issue}</li>
              ))}
            </ul>
          )
        ) : null}
        {reviewDialogMode === "performance" && workbookReviewStats ? (
          <div
            data-testid="spreadsheet-review-performance-summary"
            style={{ display: "grid", gap: 8, fontSize: 12 }}
          >
            <strong>
              {workbookReviewStats.cellsWithData > 50000 ||
              workbookReviewStats.objects > 200 ||
              workbookReviewStats.conditionalFormats > 500
                ? t("spreadsheet.review.performanceWarning")
                : t("spreadsheet.review.performanceHealthy")}
            </strong>
            <span>
              {t("spreadsheet.review.cellsWithData")}:{" "}
              {workbookReviewStats.cellsWithData}
            </span>
            <span>
              {t("spreadsheet.review.formulas")}: {workbookReviewStats.formulas}
            </span>
            <span>
              {t("spreadsheet.review.objects")}: {workbookReviewStats.objects}
            </span>
            <span>
              {t("spreadsheet.review.conditionalFormats")}:{" "}
              {workbookReviewStats.conditionalFormats}
            </span>
          </div>
        ) : null}
        {reviewDialogMode === "statistics" && workbookReviewStats ? (
          <dl
            data-testid="spreadsheet-review-statistics-grid"
            style={{
              display: "grid",
              gridTemplateColumns: "minmax(150px, 1fr) auto",
              gap: "7px 18px",
              margin: 0,
              fontSize: 12,
            }}
          >
            {(
              [
                ["sheets", workbookReviewStats.sheets],
                ["cellsWithData", workbookReviewStats.cellsWithData],
                ["formulas", workbookReviewStats.formulas],
                ["tables", workbookReviewStats.tables],
                ["pivots", workbookReviewStats.pivots],
                ["charts", workbookReviewStats.charts],
                ["images", workbookReviewStats.images],
                ["objects", workbookReviewStats.objects],
                ["comments", workbookReviewStats.comments],
                ["mergedRanges", workbookReviewStats.mergedRanges],
                [
                  "conditionalFormats",
                  workbookReviewStats.conditionalFormats,
                ],
              ] as const
            ).map(([key, value]) => (
              <Fragment key={key}>
                <dt>{t(`spreadsheet.review.${key}`)}</dt>
                <dd style={{ margin: 0, fontVariantNumeric: "tabular-nums" }}>
                  {value}
                </dd>
              </Fragment>
            ))}
          </dl>
        ) : null}
      </OfficeDialogShell>
  );
}
