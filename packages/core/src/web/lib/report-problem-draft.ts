/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * "Report a problem" from inside an editor: the notebook's Report a problem
 * panel opened with a draft the editor wrote (what failed, where, why). The
 * member reads, edits and sends it themselves -- nothing is sent from here.
 *
 * One window event: the notebook rail answers it by opening its panel with
 * the draft and marking it handled. `openReportProblemDraft` says whether a
 * panel took it, so the editor can fall back (copy the words) where no rail
 * is mounted (a guest, an embed).
 */

export const REPORT_PROBLEM_DRAFT_EVENT = "mrai:report-problem-draft";

export interface ReportProblemDraftDetail {
  text: string;
  handled: boolean;
}

export function openReportProblemDraft(text: string): boolean {
  if (typeof window === "undefined") return false;
  const detail: ReportProblemDraftDetail = { text, handled: false };
  window.dispatchEvent(new CustomEvent<ReportProblemDraftDetail>(REPORT_PROBLEM_DRAFT_EVENT, { detail }));
  return detail.handled;
}
