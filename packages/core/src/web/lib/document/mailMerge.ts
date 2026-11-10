/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word-compatible Finish & Merge record planning.
 *
 * The established mail-merge model remains the authority for recipient
 * normalization, filtering, sorting and field substitution. This module adds
 * the final-output selection boundary used by both the dialog and DOCX export:
 * All, the currently previewed record, or an inclusive 1-based From/To window
 * over the filtered/sorted merge order.
 */

import {
  selectMailMergeRecords,
  type MailMergeData,
} from "../../components/console/editor/native-document/mailMerge";

export * from "../../components/console/editor/native-document/mailMerge";

/** A browser-local merge is intentionally bounded before DOCX allocation. */
export const MAX_MAIL_MERGE_FINISH_RECORDS = 500;

export type MailMergeFinishSelection =
  | { mode: "all" }
  | { mode: "current"; currentRecordNumber: number }
  | { mode: "range"; from: number; to: number };

export type MailMergeFinishPlanStatus =
  | "ready"
  | "empty"
  | "invalid"
  | "over-limit";

export interface MailMergeFinishPlan {
  status: MailMergeFinishPlanStatus;
  /** Original 1-based source rows, in the filtered/sorted output order. */
  recordNumbers: readonly number[];
  /** Records remaining after the stored filter/sort/skip rules. */
  eligibleCount: number;
  /** Size of a valid request before the local safety cap is applied. */
  candidateCount: number;
}

function planResult(
  status: MailMergeFinishPlanStatus,
  eligibleCount: number,
  candidateRecordNumbers: readonly number[] = [],
): MailMergeFinishPlan {
  return {
    status,
    recordNumbers: status === "ready" ? candidateRecordNumbers : [],
    eligibleCount,
    candidateCount: candidateRecordNumbers.length,
  };
}

function isPositiveRecordNumber(value: number): boolean {
  return Number.isInteger(value) && value > 0;
}

/**
 * Resolve one Finish & Merge choice without mutating the template or recipient
 * store. Invalid/zero/out-of-bounds ranges deliberately resolve to zero output;
 * they are never clamped or widened to All.
 */
export function planMailMergeFinish(
  data: MailMergeData,
  selection: MailMergeFinishSelection,
): MailMergeFinishPlan {
  const records = selectMailMergeRecords(data);
  const eligibleCount = records.length;
  if (eligibleCount === 0) return planResult("empty", 0);

  let candidates: readonly number[];
  switch (selection.mode) {
    case "all":
      candidates = records.map((record) => record.number);
      break;
    case "current": {
      if (!isPositiveRecordNumber(selection.currentRecordNumber)) {
        return planResult("invalid", eligibleCount);
      }
      const current = records.find(
        (record) => record.number === selection.currentRecordNumber,
      );
      if (!current) return planResult("invalid", eligibleCount);
      candidates = [current.number];
      break;
    }
    case "range": {
      const { from, to } = selection;
      if (
        !isPositiveRecordNumber(from) ||
        !isPositiveRecordNumber(to) ||
        from > to ||
        to > eligibleCount
      ) {
        return planResult("invalid", eligibleCount);
      }
      candidates = records
        .slice(from - 1, to)
        .map((record) => record.number);
      break;
    }
  }

  if (candidates.length === 0) return planResult("empty", eligibleCount);
  if (candidates.length > MAX_MAIL_MERGE_FINISH_RECORDS) {
    return planResult("over-limit", eligibleCount, candidates);
  }
  return planResult("ready", eligibleCount, candidates);
}
