/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a person saw go wrong with an Office file, as a code and a format.
 *
 * The server's ledger says what it did; this says what was on the screen: a
 * save the window says failed, a partial write-back in the status bar. It goes
 * through the client error beacon into `client_error_reports`, read by the
 * 监控 page and the alert mail (`app/services/writeback_experience.py`). Only
 * an allow-listed format and a stable code leave the browser, never text,
 * titles or ids of what was on the page. Each thing is reported once per page
 * in ten minutes, and a browser with Do-Not-Track sends nothing.
 */
import { isDoNotTrackEnabled, sendErrorBeacon } from "./errors/report-beacon";

export type OfficeExperienceKind = "save_failure" | "partial" | "smartart_kept" | "pdf_long_way";

const NAMES: Readonly<Record<OfficeExperienceKind, string>> = {
  save_failure: "OfficeSaveFailureShown",
  partial: "OfficePartialShown",
  // A SmartArt diagram kept as Office drew it instead of converting, by why
  // (`drawing-mismatch`, `definition-unavailable`, ...: SMARTART-GAP.md §8.6).
  smartart_kept: "OfficeSmartArtKept",
  // A PDF revision that could not go up in one tick and was read in the long
  // way, its boxes out meanwhile, by why. The layers kept for that way (the
  // queued open, its adoption, the ghost) go once this is zero on the live site.
  pdf_long_way: "PdfRevisionLongWay",
};

const FORMATS: Readonly<Record<string, string>> = {
  presentation: "pptx",
  spreadsheet: "xlsx",
  document: "docx",
  pdf: "pdf",
};

const CODE = /^[A-Za-z0-9_.:-]{1,96}$/u;
const REPEAT_AFTER_MS = 10 * 60 * 1000;
const lastReported = new Map<string, number>();

/** Returns whether a report went out. Never throws. */
export function reportOfficeExperience(
  kind: OfficeExperienceKind,
  input: { pageId: string; pageType: string | null | undefined; code: string | null | undefined },
  now = Date.now(),
): boolean {
  try {
    const format = input.pageType ? FORMATS[input.pageType] : undefined;
    if (!format || typeof window === "undefined" || isDoNotTrackEnabled()) return false;
    const code = input.code && CODE.test(input.code) ? input.code : "unknown";
    const key = `${kind}|${input.pageId}|${code}`;
    const last = lastReported.get(key);
    if (last !== undefined && now - last < REPEAT_AFTER_MS) return false;
    lastReported.set(key, now);
    if (lastReported.size > 500) lastReported.clear();
    sendErrorBeacon({
      name: NAMES[kind],
      message: code,
      digest: code,
      // The kind of window, not which page: the path holds no page's id here.
      path: "/office",
      source: "window",
      window_type: format,
    });
    return true;
  } catch {
    return false;
  }
}

/** Test seam. */
export function resetOfficeExperienceReportsForTests(): void {
  lastReported.clear();
}
