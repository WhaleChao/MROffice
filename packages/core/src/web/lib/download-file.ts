/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { browserAppAliasOrigin, sisterFileOrigin } from "./app-alias-origin";
import { checkFileAccess, fileAccessRepairApplies } from "./file-access-repair";

let downloadGate: (() => void) | null = null;

/**
 * While a visitor's editor is open (the guest notebook's sample PDF, whose
 * edits never leave the page), every download asks for an account through
 * `gate` instead of producing a file; returns the release.
 */
export function holdDownloadsBehind(gate: () => void): () => void {
  downloadGate = gate;
  return () => {
    if (downloadGate === gate) downloadGate = null;
  };
}

/** True when the download was answered by the gate instead. */
function downloadWasGated(): boolean {
  if (!downloadGate) return false;
  downloadGate();
  return true;
}

/**
 * Trigger a client-side download of in-memory text as a file.
 *
 * Reuse the Blob download lifecycle so every caller keeps the object URL alive
 * through the browser's click/download task. The boolean lets a visible command
 * report an unavailable browser capability instead of silently doing nothing.
 */
export function downloadTextFile(
  filename: string,
  text: string,
  mimeType = "text/plain;charset=utf-8",
): boolean {
  if (typeof Blob === "undefined") return false;
  return downloadBlobFile(filename, new Blob([text], { type: mimeType }));
}

/**
 * Raised (on `window`) for a download that is ready only after its click's
 * user activation has expired. The shell's DownloadReadyGuard offers a Save
 * button and sets `detail.handled`.
 */
export const DOWNLOAD_NEEDS_GESTURE_EVENT = "mrai:download-needs-gesture";

export interface DownloadNeedsGestureDetail {
  readonly name: string;
  readonly start: () => void;
  /** "open": the file opens in a new tab (a PDF to print) rather than
   * saving; the notice then offers Open. Absent means "save". */
  readonly action?: "save" | "open";
  handled: boolean;
}

/**
 * Ask the shell for a tap that opens a tab: a window.open after a long await
 * (a PDF rendered for printing) is refused as a pop-up, on iPad Safari always.
 * True when the shell took the offer.
 */
export function offerOpenWhenAllowed(name: string, open: () => void): boolean {
  if (typeof window === "undefined") return false;
  const detail: DownloadNeedsGestureDetail = { name, start: open, action: "open", handled: false };
  window.dispatchEvent(new CustomEvent(DOWNLOAD_NEEDS_GESTURE_EVENT, { detail }));
  return detail.handled;
}

/**
 * Start a download that may have finished long after the click that asked
 * for it.
 *
 * A browser lets a page start a download without a fresh click only so
 * often: Chrome blocks one more "automatic" download behind a small prompt
 * in the address bar, so a workbook export that took minutes ended with the
 * export reported ready, no file and no message (user audit S-098). The
 * download always starts at once, as before; when the click's activation has
 * run out, the shell also offers a Save button for the case the browser held
 * it back, a gesture every browser honours. Nothing is asked when the click
 * still counts.
 */
export function startDownloadWhenAllowed(name: string, start: () => boolean): boolean {
  const started = start();
  const activation =
    typeof navigator === "undefined"
      ? undefined
      : (navigator as Navigator & { userActivation?: { isActive: boolean } }).userActivation;
  if (!started || !activation || activation.isActive || typeof window === "undefined") return started;
  const detail: DownloadNeedsGestureDetail = {
    name,
    start: () => {
      start();
    },
    handled: false,
  };
  window.dispatchEvent(new CustomEvent(DOWNLOAD_NEEDS_GESTURE_EVENT, { detail }));
  return started;
}

/** Whether this realm can start an in-memory browser download. */
export function isBrowserDownloadAvailable(): boolean {
  return (
    typeof document !== "undefined" &&
    typeof Blob !== "undefined" &&
    typeof URL !== "undefined" &&
    typeof URL.createObjectURL === "function"
  );
}

/**
 * Trigger a browser download for an already-built Blob. Keeping the transient
 * object-URL lifecycle in one place prevents server-backed export callers from
 * accidentally opening authenticated API responses in a new tab.
 */
export function downloadBlobFile(filename: string, blob: Blob): boolean {
  if (downloadWasGated()) return true;
  if (!isBrowserDownloadAvailable()) return false;
  const url = URL.createObjectURL(blob);
  let anchor: HTMLAnchorElement | null = null;
  try {
    anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.rel = "noopener";
    document.body.appendChild(anchor);
    anchor.click();
    return true;
  } finally {
    try {
      if (anchor) {
        try {
          anchor.remove();
        } catch (removeError) {
          try {
            anchor.parentNode?.removeChild(anchor);
          } finally {
            throw removeError;
          }
        }
      }
    } finally {
      // Keep the object URL alive through the browser's click/download task.
      // Immediate revocation is racy in Firefox/WebKit for larger responses.
      try {
        // WebKit can defer a large authenticated export well beyond the click
        // task. A bounded grace period keeps DOCX/PDF downloads alive without
        // leaking the URL for the lifetime of the page.
        setTimeout(() => URL.revokeObjectURL(url), 60_000);
      } catch {
        URL.revokeObjectURL(url);
      }
    }
  }
}

/**
 * Raised (on `window`) when a server-issued download URL answered with a page
 * instead of a file (the file gate's refusal, an expired link): the shell's
 * DownloadReadyGuard says so.
 */
export const DOWNLOAD_FAILED_EVENT = "mrai:download-failed";

export interface DownloadFailedDetail {
  readonly name: string;
  /** Set when another notice already says why (the file session's
   * FileAccessGuard, the site's AccessGateGuard). */
  readonly reason?: "file-access" | "access-gate";
}

/** How long a download frame is kept: removing it can cancel the transfer. */
const DOWNLOAD_FRAME_LIFETIME_MS = 10 * 60_000;

function openFileHostDownloadFrame(filename: string, href: string, retried: boolean): void {
  const frame = document.createElement("iframe");
  frame.hidden = true;
  frame.setAttribute("aria-hidden", "true");
  frame.tabIndex = -1;
  frame.referrerPolicy = "no-referrer";
  frame.style.display = "none";
  const removal = window.setTimeout(() => frame.remove(), DOWNLOAD_FRAME_LIFETIME_MS);
  frame.addEventListener("load", () => {
    window.clearTimeout(removal);
    frame.remove();
    void fileHostDownloadRefused(filename, href, retried);
  }, { once: true });
  frame.src = href;
  document.body.appendChild(frame);
}

/**
 * The file host answered a download frame with a page. A frame's page
 * cannot be read across origins, so ask the file host whether this
 * browser's file session is the reason (lib/file-access-repair.ts): when it
 * was and is reconnected now, the same short-lived link is opened once
 * more; otherwise the failure is reported.
 */
async function fileHostDownloadRefused(filename: string, href: string, retried: boolean): Promise<void> {
  let reason: DownloadFailedDetail["reason"];
  if (!retried && fileAccessRepairApplies()) {
    const outcome = await checkFileAccess().catch(() => "unreachable" as const);
    if (outcome === "repaired") {
      openFileHostDownloadFrame(filename, href, true);
      return;
    }
    if (outcome === "gate-locked") reason = "access-gate";
    else if (outcome === "guarded" || outcome === "refused" || outcome === "handoff-unavailable") {
      reason = "file-access";
    }
  }
  const detail: DownloadFailedDetail = reason ? { name: filename, reason } : { name: filename };
  window.dispatchEvent(new CustomEvent(DOWNLOAD_FAILED_EVENT, { detail }));
}

/** The private file host (the app's frame-src names it, proxy.ts), when this
 * URL is on it and it is not the app's own origin. */
function isFileHostUrl(url: URL): boolean {
  if (typeof window === "undefined" || url.origin === window.location.origin) return false;
  try {
    const configured = process.env.NEXT_PUBLIC_ASSET_ORIGIN;
    if (!configured) return false;
    // A sister site's pages use that site's own file host.
    const own = sisterFileOrigin(browserAppAliasOrigin()) ?? new URL(configured).origin;
    return own === url.origin;
  } catch {
    return false;
  }
}


















export function downloadUrlFile(filename: string, rawUrl: string): boolean {
  if (downloadWasGated()) return true;
  if (typeof document === "undefined") return false;

  let url: URL;
  try {
    if (rawUrl.length > 16_384) return false;
    url = new URL(rawUrl, document.baseURI);
  } catch {
    return false;
  }
  if (
    (url.protocol !== "https:" && url.protocol !== "http:") ||
    url.username !== "" ||
    url.password !== ""
  ) {
    return false;
  }

  if (isFileHostUrl(url)) {
    openFileHostDownloadFrame(filename, url.href, false);
    return true;
  }

  const anchor = document.createElement("a");
  try {
    anchor.href = url.href;
    anchor.download = filename;
    anchor.rel = "noopener noreferrer";
    anchor.referrerPolicy = "no-referrer";
    document.body.appendChild(anchor);
    anchor.click();
    return true;
  } finally {
    anchor.remove();
  }
}
