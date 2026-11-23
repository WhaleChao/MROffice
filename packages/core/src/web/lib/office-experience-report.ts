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

export type OfficeExperienceKind = "save_failure" | "partial" | "smartart_kept" | "pdf_long_way" | "layout_twitch"
  | "room_durable_stalled" | "save_surface_stalled";

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
  // A document's pages went somewhere and came back, or a guard had to stop
  // two layouts alternating (paged/ledger/ledgerTwitch), by which: `returned`,
  // `held-cycle`, `held-repair`. The reader saw the page jump; the bar is zero.
  layout_twitch: "DocsLayoutTwitch",
  // A deck room's updates the database has not confirmed for a minute while
  // its socket is up (lib/save-stall-report): the window says "Save" though
  // nothing is left to send (S6 acc4, a reconnect that lost the waiting ids).
  room_durable_stalled: "PageRoomDurableStalled",
  // A save face saying "Saving..." or "Save" for 90 s with no request on its
  // way, by the inputs that make it so (lib/save-stall-report).
  save_surface_stalled: "SaveSurfaceStalled",
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

/** One declined change as a code pair: the control (`object:text`,
 *  `background`, `slideEdits`, `deck:<name>`...) and the server's reason code
 *  (`not_yet`, `readback_mismatch`...). Never a word of the page. */
export interface OfficeExperienceDeclined {
  readonly control?: unknown;
  readonly reason?: unknown;
}

/** Where the 监控 page can find what was shown without the page's content: a
 *  hashed page id (FNV-1a 64, hex) matchable against the server's own page
 *  ids, the declined controls with their reason codes, and a failed write's
 *  reason code (its first two segments: `pptx_patchback_rejected:text_not_patchable`). */
export interface OfficeExperienceDetail {
  readonly declined?: readonly OfficeExperienceDeclined[] | null;
  readonly failure?: unknown;
}

const MAX_DECLINED_CODES = 8;

/** FNV-1a 64 of ``value``, 16 hex digits: a page's reference in a report. */
export function officeExperiencePageRef(value: string): string {
  let hash = 0xcbf29ce484222325n;
  for (const byte of new TextEncoder().encode(value)) {
    hash ^= BigInt(byte);
    hash = (hash * 0x100000001b3n) & 0xffffffffffffffffn;
  }
  return hash.toString(16).padStart(16, "0");
}

function codeOf(value: unknown, segments?: number): string | null {
  if (typeof value !== "string") return null;
  const trimmed = (segments ? value.split(":").slice(0, segments).join(":") : value).trim();
  return CODE.test(trimmed) ? trimmed : null;
}

/** The report's message: the code, then the page's reference and what was
 *  declined or why the write failed, as codes only. */
export function officeExperienceMessage(code: string, pageId: string, detail?: OfficeExperienceDetail): string {
  const parts = [code, `page=${officeExperiencePageRef(pageId)}`];
  const declined: string[] = [];
  for (const entry of detail?.declined ?? []) {
    if (!entry || typeof entry !== "object") continue;
    const control = codeOf(entry.control);
    if (!control) continue;
    const reason = codeOf(entry.reason);
    const item = reason ? `${control}/${reason}` : control;
    if (!declined.includes(item)) declined.push(item);
    if (declined.length >= MAX_DECLINED_CODES) break;
  }
  if (declined.length) parts.push(`declined=${declined.join(",")}`);
  const failure = codeOf(detail?.failure, 2);
  if (failure) parts.push(`failure=${failure}`);
  return parts.join(" ");
}

/** Returns whether a report went out. Never throws. */
export function reportOfficeExperience(
  kind: OfficeExperienceKind,
  input: {
    pageId: string;
    pageType: string | null | undefined;
    code: string | null | undefined;
    detail?: OfficeExperienceDetail;
  },
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
      // The digest stays the bare code (the 监控 page counts by it); the
      // message adds the page's hashed reference and the declined codes.
      message: input.detail ? officeExperienceMessage(code, input.pageId, input.detail) : code,
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
const MAX_HIDDEN_ISSUE_CODES = 12;







export function reportHiddenDocumentIssues(
  input: {
    pageId: string | null | undefined;
    surface: string | null | undefined;
    keys: readonly string[];
  },
  now = Date.now(),
): boolean {
  try {
    if (typeof window === "undefined" || isDoNotTrackEnabled()) return false;
    const codes: string[] = [];
    for (const key of input.keys) {
      const code = codeOf(key) ?? "unknown";
      if (!codes.includes(code)) codes.push(code);
      if (codes.length >= MAX_HIDDEN_ISSUE_CODES) break;
    }
    if (!codes.length) return false;
    const surface = codeOf(input.surface) ?? "unknown";
    const page = input.pageId ? officeExperiencePageRef(input.pageId) : "none";
    const key = `hidden|${page}|${surface}|${codes.join(",")}`;
    const last = lastReported.get(key);
    if (last !== undefined && now - last < REPEAT_AFTER_MS) return false;
    lastReported.set(key, now);
    if (lastReported.size > 500) lastReported.clear();
    sendErrorBeacon({
      name: "OfficeIssueHidden",
      message: `issues=${codes.join(",")} surface=${surface} page=${page}`,
      digest: codes[0],
      path: "/office",
      source: "window",
      window_type: "office",
    });
    return true;
  } catch {
    return false;
  }
}

export function resetOfficeExperienceReportsForTests(): void {
  lastReported.clear();
}
