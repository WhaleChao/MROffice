/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Client error-report beacon transport.
 *
 * The production observability side-channel for the React error boundaries
 * (`app/[locale]/error.tsx` + `app/global-error.tsx`). When an unhandled render
 * error trips a boundary, the boundary renders a friendly, localized recovery
 * UI for the user AND fires this fire-and-forget beacon so the crash is
 * *observable* in the backend structured-log stream instead of dying in the
 * user's console alone:
 *
 *     POST /api/v1/telemetry/error   { message, name, digest, path, stack }
 *
 * It deliberately mirrors `lib/web-vitals/beacon.ts`:
 *   - `navigator.sendBeacon` first (survives the page-unload that often
 *     follows a crash), falling back to `fetch(..., { keepalive: true })`.
 *   - Framework-free + structurally-typed input so it is trivially unit-tested
 *     with a plain object (not coupled to React's `Error & { digest }`).
 *
 * Privacy / safety contract (mirrors the backend route's red-lines):
 *   - Only `pathname` is sent for the location — never the query string or
 *     fragment, so a URL carrying `?token=` / `?email=` can't leak.
 *   - `message` + `stack` are length-capped *client-side* before they leave the
 *     browser (the backend re-caps + PII-value-scrubs them as defence in
 *     depth). This payload is NEVER shown to the user — the boundary UI only
 *     renders generic localized copy.
 *   - Respects Do-Not-Track: when the browser signals DNT, nothing is sent.
 *   - Fire-and-forget: any transport failure is swallowed (never throws back
 *     into the error boundary, which is already handling an error).
 */
import { getApiHttpBaseUrl } from "../env";

/** Path on the API the beacon POSTs to. */
export const ERROR_BEACON_PATH = "/api/v1/telemetry/error";

/**
 * Client-side caps. The backend caps again (these need only bound what leaves
 * the browser). The message cap matches a short human sentence; the stack cap
 * keeps a useful handful of frames without shipping a giant trace.
 */
export const MAX_REPORTED_MESSAGE_LEN = 1024;
export const MAX_REPORTED_STACK_LEN = 8192;

/**
 * The compact beacon body. Field names are 1-to-1 with the backend
 * `ErrorReportBeacon` schema.
 */
export interface ErrorReportBeacon {
  /** Error message (truncated). */
  message?: string;
  /** Error constructor name, e.g. "TypeError". */
  name?: string;
  /** Next.js error digest (server-minted hash; not PII). */
  digest?: string;
  /** Page pathname the error occurred on (no query string / fragment). */
  path: string;
  /** Stack trace (truncated). Omitted in production unless allowed (see below). */
  stack?: string;
  /** Which boundary caught it: a route, the root layout, or one window. */
  source?: ErrorReportSource;
  /** The window's type when `source` is "window" (e.g. "note", "file"). */
  window_type?: string;
}

/** The boundaries that report. */
export type ErrorReportSource = "route" | "global" | "window";

/** Where a report came from, as the boundary that caught it knows. */
export interface ErrorReportContext {
  source?: ErrorReportSource;
  /** The window's type; a short token, dropped when it is not one. */
  windowType?: string;
}









export const MAX_WINDOW_STACK_LINES = 8;

/**
 * The structural subset of `Error & { digest? }` this module consumes. Kept
 * structural (not `Error`) so the helper is testable with a plain object and
 * tolerant of non-Error throws.
 */
export interface ErrorLike {
  message?: unknown;
  name?: unknown;
  stack?: unknown;
  digest?: unknown;
}

/**
 * What the public entry points accept. `catch (e)` is typed `unknown`, and a
 * throw can be any value (a string, a plain object, `null`), so we take the
 * widest type and narrow defensively inside. `ErrorLike` is the *shape* we
 * read; `ReportableError` is the *accepted* parameter type.
 */
export type ReportableError = unknown;

/** Read an `ErrorLike` view of an arbitrary thrown value (object or not). */
function asErrorLike(error: ReportableError): ErrorLike {
  return error && typeof error === "object" ? (error as ErrorLike) : {};
}

/**
 * Return true when the browser signals Do-Not-Track. Mirrors the web-vitals
 * beacon: checks the three historic surfaces; any `"1"` / `"yes"` means opt-out.
 */
export function isDoNotTrackEnabled(): boolean {
  if (typeof navigator === "undefined" && typeof window === "undefined") {
    return false;
  }
  const signals: Array<string | null | undefined> = [
    typeof navigator !== "undefined" ? navigator.doNotTrack : undefined,
    typeof window !== "undefined"
      ? (window as unknown as { doNotTrack?: string }).doNotTrack
      : undefined,
    typeof navigator !== "undefined"
      ? (navigator as unknown as { msDoNotTrack?: string }).msDoNotTrack
      : undefined,
  ];
  return signals.some((v) => v === "1" || v === "yes");
}

/** Current pathname only — deliberately drops `search` + `hash` (privacy). */
function currentPath(): string {
  if (typeof window === "undefined" || !window.location) {
    return "/";
  }
  return window.location.pathname || "/";
}

/** Absolute (or same-origin) beacon URL, matching how `lib/api.ts` builds it. */
function beaconUrl(): string {
  return `${getApiHttpBaseUrl()}${ERROR_BEACON_PATH}`;
}

/** Coerce an unknown field to a trimmed, length-capped string (or undefined). */
function asCappedString(value: unknown, limit: number): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  return trimmed.slice(0, limit);
}

/**
 * True in a production build. In production we DROP the stack from the beacon
 * by default: a minified production stack carries little triage value (the
 * `digest` correlates to the server-side build map instead), and not sending it
 * is the strictest privacy posture. In development we keep the (truncated)
 * stack because it is genuinely useful at the dev console / local telemetry.
 *
 * `NEXT_PUBLIC_ENABLE_ERROR_STACK_REPORT=1` opts a production deployment back
 * into shipping the truncated stack (the backend still PII-scrubs + re-caps it),
 * for teams that wire a real source-map-aware sink.
 */
function shouldReportStack(): boolean {
  if (process.env.NODE_ENV !== "production") return true;
  return process.env.NEXT_PUBLIC_ENABLE_ERROR_STACK_REPORT === "1";
}

/**
 * Build the compact beacon body from an error-like object. The `path` is the
 * privacy-safe pathname; `message` + `stack` are trimmed + capped; `stack` is
 * dropped in production by default (see {@link shouldReportStack}).
 */
export function buildErrorBeacon(
  error: ReportableError,
  context: ErrorReportContext = {},
): ErrorReportBeacon {
  const view = asErrorLike(error);
  const message = asCappedString(view.message, MAX_REPORTED_MESSAGE_LEN);
  const name = asCappedString(view.name, 256);
  const digest = asCappedString(view.digest, 256);
  const stack = shouldReportStack()
    ? asCappedString(view.stack, MAX_REPORTED_STACK_LEN)
    : context.source === "window"
      ? asCappedString(
          typeof view.stack === "string"
            ? view.stack.split("\n").slice(0, MAX_WINDOW_STACK_LINES + 1).join("\n")
            : undefined,
          MAX_REPORTED_STACK_LEN,
        )
      : undefined;
  const windowType =
    context.source === "window" && typeof context.windowType === "string" &&
    /^[a-z0-9_.-]{1,40}$/.test(context.windowType)
      ? context.windowType
      : undefined;
  return {
    ...(message ? { message } : {}),
    ...(name ? { name } : {}),
    ...(digest ? { digest } : {}),
    ...(stack ? { stack } : {}),
    ...(context.source ? { source: context.source } : {}),
    ...(windowType ? { window_type: windowType } : {}),
    path: currentPath(),
  };
}

/** Reports already sent from this page, so a window that fails on every
 *  render (or every Retry) is reported once, not once per paint. */
const reportedKeys = new Set<string>();
const MAX_REPORTED_KEYS = 50;

function firstReport(body: ErrorReportBeacon): boolean {
  if (body.source !== "window") return true;
  const key = [body.window_type, body.name, body.message].join("|");
  if (reportedKeys.has(key) || reportedKeys.size >= MAX_REPORTED_KEYS) return false;
  reportedKeys.add(key);
  return true;
}

/** Test seam: forget the reports already sent from this page. */
export function resetReportedErrorsForTests(): void {
  reportedKeys.clear();
}

/**
 * Ship the beacon. Prefers `navigator.sendBeacon` (survives the page-unload
 * that frequently follows a crash); falls back to `fetch(keepalive)`. Returns
 * true if a transport was dispatched. Never throws.
 */
export function sendErrorBeacon(body: ErrorReportBeacon): boolean {
  const url = beaconUrl();
  const payload = JSON.stringify(body);

  try {
    if (
      typeof navigator !== "undefined" &&
      typeof navigator.sendBeacon === "function"
    ) {
      const blob = new Blob([payload], { type: "application/json" });
      if (navigator.sendBeacon(url, blob)) {
        return true;
      }
    }
  } catch {
    // sendBeacon can throw (e.g. payload too large for the UA queue); fall
    // through to fetch.
  }

  try {
    if (typeof fetch === "function") {
      void fetch(url, {
        method: "POST",
        body: payload,
        headers: { "Content-Type": "application/json" },
        keepalive: true,
        credentials: "omit",
        cache: "no-store",
        mode: "cors",
      }).catch(() => {
        /* swallow — fire-and-forget */
      });
      return true;
    }
  } catch {
    /* swallow — fire-and-forget */
  }

  return false;
}

/**
 * The single entry point the error boundaries call. Honors DNT, builds the
 * privacy-safe body, and ships it. Returns true when a beacon was dispatched
 * (false when suppressed by DNT or no transport). Never throws — a failure here
 * must not compound the error the boundary is already handling.
 */
/** A stack frame that runs code the page did not ship: a browser extension
 * or a user script injected into it. */
const THIRD_PARTY_FRAME = /(?:^|[\s@(])(?:user-script:|chrome-extension:\/\/|moz-extension:\/\/|safari(?:-web)?-extension:\/\/|webkit-masked-url:\/\/)/;








export function thrownByThirdPartyScript(error: ReportableError): boolean {
  const stack = error && typeof error === "object" ? (error as ErrorLike).stack : undefined;
  if (typeof stack !== "string") return false;
  const frames = stack
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line && (line.startsWith("at ") || line.includes("@")));
  // Frames the engine itself adds ("getPrototypeOf@[native code]") say nothing.
  const first = frames.find((line) => !line.includes("[native code]"));
  return Boolean(first && THIRD_PARTY_FRAME.test(first));
}

export function reportClientError(
  error: ReportableError,
  context?: ErrorReportContext,
): boolean {
  try {
    if (isDoNotTrackEnabled()) {
      return false;
    }
    if (thrownByThirdPartyScript(error)) {
      return false;
    }
    const body = buildErrorBeacon(error, context);
    // Nothing meaningful to report (no message / name / stack) — skip.
    if (!body.message && !body.name && !body.stack) {
      return false;
    }
    if (!firstReport(body)) {
      return false;
    }
    return sendErrorBeacon(body);
  } catch {
    return false;
  }
}
