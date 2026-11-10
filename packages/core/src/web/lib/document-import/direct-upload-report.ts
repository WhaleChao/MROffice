/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How an upload the server offered the direct road to went: every part
 * straight to object storage, or some through the upload proxy route, and
 * why the first one did. It goes through the client error beacon into
 * `client_error_reports`, whose row carries the browser engine the API read
 * from the beacon's User-Agent, so the share of uploads that fall back can be
 * read there by engine (Safari's WebKit above all) without anyone uploading a
 * file to look:
 *
 *   name `DirectUploadSent`: every part went straight to storage;
 *   name `DirectUploadFellBack`: at least one part took the proxy route, and
 *   `digest` says why the first did (a stable code from {@link DirectUploadFallback}).
 *
 * They are not errors: the API logs them at info under their own event and
 * leaves them out of every count of client errors (`OUTCOME_REPORTS` in
 * `app/services/observability/client_error_store.py`; add a new name there).
 *
 * Only the code and the part counts leave the browser, never the file's name,
 * size, address or any id. One report per upload; a browser with
 * Do-Not-Track sends nothing.
 */
import { isDoNotTrackEnabled, sendErrorBeacon } from "../errors/report-beacon";

/** Why a part that could go straight to storage took the proxy route. */
export type DirectUploadFallback =
  /** An earlier upload of this page could not reach storage at all. */
  | "storage_unreachable_earlier"
  /** The part addresses were not given (the server withdrew the road, or a request failed). */
  | `part_urls_${string}`
  /** The PUT never reached storage: a network, TLS, CORS or mixed-content failure. */
  | "storage_unreachable"
  /** Storage answered the PUT with this status. */
  | `storage_${number}`
  /** None of the above was noted (never expected). */
  | "other";

export const DIRECT_UPLOAD_SENT = "DirectUploadSent";
export const DIRECT_UPLOAD_FELL_BACK = "DirectUploadFellBack";

const CODE = /^[a-z0-9_]{1,64}$/u;

/** Returns whether a report went out. Never throws. */
export function reportDirectUploadOutcome(outcome: {
  fallback: DirectUploadFallback | null;
  direct: number;
  proxied: number;
}): boolean {
  try {
    if (typeof window === "undefined" || isDoNotTrackEnabled()) return false;
    const code = outcome.fallback && CODE.test(outcome.fallback) ? outcome.fallback : outcome.fallback ? "other" : null;
    return sendErrorBeacon({
      name: code ? DIRECT_UPLOAD_FELL_BACK : DIRECT_UPLOAD_SENT,
      message: `${code ?? "direct"} direct=${outcome.direct} proxied=${outcome.proxied}`,
      digest: code ?? "direct",
      path: "/upload",
      source: "window",
      window_type: "upload",
    });
  } catch {
    return false;
  }
}
