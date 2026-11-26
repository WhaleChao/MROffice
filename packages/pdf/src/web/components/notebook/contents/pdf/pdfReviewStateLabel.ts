/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback } from "react";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import type { PdfAnnotationReviewState } from "@mrnote/core/web/lib/pdf/annotation-replies.ts";

/** The review states with a label of their own (ISO 32000-1 Table 172). */
export const PDF_REVIEW_STATE_NAMES = ["Accepted", "Rejected", "Cancelled", "Completed", "None", "Marked", "Unmarked"];

type Translator = ReturnType<typeof useTranslations<"console-notebooks">>;

/** One wording of a review state for every comment surface of a PDF. */
export function pdfReviewStateLabel(t: Translator, state: PdfAnnotationReviewState | undefined): string {
  if (!state) return "";
  const key = PDF_REVIEW_STATE_NAMES.includes(state.state) ? state.state : "Other";
  return t(`pdf.reviewState.${key}` as Parameters<Translator>[0], { state: state.state });
}

export function usePdfReviewStateLabel(t: Translator): (state: PdfAnnotationReviewState | undefined) => string {
  return useCallback((state) => pdfReviewStateLabel(t, state), [t]);
}
