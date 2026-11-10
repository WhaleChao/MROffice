/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  PageExportFidelityRead,
  PageExportNoticeRead,
} from "./notebook-sdk";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type TranslateFn = (key: any, values?: any) => string;

/** Notice codes the locale files phrase; anything else uses the server title. */
const PHRASED = new Set([
  "page_count_differs",
  "header_footer_differs",
  "content_clipped",
  "font_fallback",
  "formatting_differs",
  "fonts_not_embedded",
  "picture_count_differs",
  "page_break_lost",
  "formatting_unproven",
  "fonts_unproven",
  "layout_unproven",
  "header_footer_unproven",
  "pictures_unproven",
  
  // the checks observe; only lost content refuses).
  "objects_overlap",
  "text_may_overflow",
  "export_keeps_less",
  "formula_reference_invalid",
  "dropdown_source_missing",
  "object_position_unproven",
  "date_weekday_mismatch",
  "requirements_unmet",
  "check_unverified",
  "structure_notice",
  // An AI edit made with Track Changes on that Compare could not redline.
  "revisions_not_recorded",
]);

export interface DocsPublicationState {
  notices: PageExportNoticeRead[];
  /** A check could not run when the document was published. */
  unverified: boolean;
  /** The editor recheck after publishing found fault: check before editing. */
  recheckFailed: boolean;
}

export function docsPublicationState(
  read: PageExportFidelityRead | null | undefined,
): DocsPublicationState {
  const notices = Array.isArray(read?.notices)
    ? read.notices.filter(
        (item) =>
          item &&
          typeof item.code === "string" &&
          (item.tier === "presentation" || item.tier === "unproven"),
      )
    : [];
  return {
    notices,
    unverified: read?.status === "unverified",
    recheckFailed: read?.editor_recheck?.status === "failed",
  };
}

function count(value: number | null | undefined): number | null {
  return typeof value === "number" && Number.isInteger(value) && value >= 0
    ? value
    : null;
}

/** One notice as a sentence in the reader's language. */
export function docsNoticeSentence(
  notice: PageExportNoticeRead,
  t: TranslateFn,
): string {
  let text = notice.title;
  if (PHRASED.has(notice.code)) {
    const expected = count(notice.expected);
    const observed = count(notice.observed);
    const counted = count(notice.count);
    if (notice.code === "page_count_differs" && expected !== null && observed !== null) {
      text = t("pages.exportNotices.code.page_count_differs_counted", { expected, observed });
    } else if (notice.code === "font_fallback" && counted !== null && counted > 0) {
      text = t("pages.exportNotices.code.font_fallback_counted", { count: counted });
    } else {
      text = t(`pages.exportNotices.code.${notice.code}`);
    }
  }
  const page = count(notice.page);
  return page !== null && page > 0
    ? t("pages.exportNotices.onPage", { page, text })
    : text;
}

/**
 * The instruction "Fix with AI" places in the page's AI panel: every notice,
 * with its page, and the one rule that matters: keep the content whole.
 */
export function docsRepairPrompt(
  notices: readonly PageExportNoticeRead[],
  t: TranslateFn,
): string {
  const items = notices
    .map((notice) => `- ${docsNoticeSentence(notice, t)}`)
    .join("\n");
  return t("pages.exportNotices.fixPrompt", { items });
}
