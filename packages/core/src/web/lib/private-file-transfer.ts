/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { browserAppAliasOrigin, sisterFileOrigin } from "./app-alias-origin";
import routes from "./private-file-routes.json";
// Function references only, resolved at call time: the repair module reads
// this module's file origin, and this module's fetch uses the repair.
import {
  noteFileAccessRequired,
  readFileGateRefusal,
  recoverFromFileGateRefusal,
  requestBodyIsReplayable,
} from "./file-access-repair";

const patterns = routes.map((route) => ({ ...route, pattern: new RegExp(route.path) }));

export function isFileTransferPath(path: string, method: string): boolean {
  return patterns.some((route) => route.methods.includes(method.toUpperCase()) && route.pattern.test(path));
}

/** The global mark of an embedded editor's frame (lib/auth-state
 * EMBED_SIGNED_IN_GLOBAL), read without importing that module. */
const EMBED_SESSION_GLOBAL = "__mroEmbedSessionSignedIn";

export function getFileTransferOrigin(): string | null {
  const raw = process.env.NEXT_PUBLIC_FILE_TRANSFER_ORIGIN?.trim();
  if (!raw) return null;
  // An embedded editor in a partner's page: the file host's gate wants this
  // site's own gate and account cookies, which a page on another site never
  // sends, and its credential is a header the file host's edge does not
  // admit. Its transfers go to the API on the page's own origin, with that
  // header (lib/embed/frame-auth), and the API hands it stored files as the
  // provider's signed links (app/services/private_download_delivery.py).
  if ((globalThis as Record<string, unknown>)[EMBED_SESSION_GLOBAL] === true) return null;
  const file = new URL(raw);
  // The file host is files.<canonical host>. A page on a configured alias
  // (cn.mr-note.com) pairs with that same canonical file host, so the
  // check is made against the canonical origin there, never the alias's own
  // files.<alias> (which does not exist).
  const main = new URL(typeof window === "undefined" || browserAppAliasOrigin()
    ? process.env.NEXT_PUBLIC_API_BASE_URL || "https://invalid.invalid"
    : window.location.origin);
  if (file.protocol !== "https:" || file.origin !== raw || file.username || file.password || file.port ||
      file.hostname !== `files.${main.hostname}`) {
    throw new Error("Private file origin is not configured safely");
  }
  // A sister site (mroffice.ai) has a file host of its own; the configured
  // one, checked above, is what makes this deployment have file hosts at all.
  return (typeof window === "undefined" ? null : sisterFileOrigin(browserAppAliasOrigin())) ?? file.origin;
}

/** Only exact owner-configured file routes may receive product credentials. */
/** The desktop app's stored files: on the page's own origin, by an opaque
 * account-bound token (apps/api app/desktop/file_delivery.py), the stand-in
 * for the site's file host. Its web server and API both want the window's
 * cookies, so these are fetched with them like the file host's files. */
const DESKTOP_FILE_PATH = /^\/api\/v1\/desktop\/files\/[A-Za-z0-9_-]{40,8192}$/;

function isDesktopFileUrl(value: string, method: string): boolean {
  if (process.env.NEXT_PUBLIC_MRAI_DESKTOP !== "classic" || typeof window === "undefined") return false;
  if (method.toUpperCase() !== "GET" && method.toUpperCase() !== "HEAD") return false;
  try {
    const url = new URL(value, window.location.origin);
    return url.origin === window.location.origin && !url.username && !url.password && !url.search &&
      DESKTOP_FILE_PATH.test(url.pathname);
  } catch {
    return false;
  }
}

export function isPrivateFileUrl(value: string, method = "GET"): boolean {
  if (isDesktopFileUrl(value, method)) return true;
  const origin = getFileTransferOrigin();
  if (!origin) return false;
  const url = new URL(value, typeof window === "undefined" ? origin : window.location.origin);
  return url.origin === origin && !url.username && !url.password && isFileTransferPath(url.pathname, method);
}

export function fileTransferApiUrl(path: string, method: string): string | null {
  if (path.startsWith("//")) return null;
  const mainOrigin = typeof window === "undefined"
    ? new URL(process.env.NEXT_PUBLIC_API_BASE_URL || "https://invalid.invalid").origin
    : window.location.origin;
  const url = new URL(path, mainOrigin);
  // On an alias page an absolute URL the API built on its canonical SITE_URL
  // names the same route, and is routed to the same file host.
  const canonicalOrigin = browserAppAliasOrigin()
    ? new URL(process.env.NEXT_PUBLIC_API_BASE_URL || "https://invalid.invalid").origin
    : mainOrigin;
  if ((url.origin !== mainOrigin && url.origin !== canonicalOrigin) || url.username || url.password || !isFileTransferPath(url.pathname, method)) return null;
  // Asked only for a file route: the print page is served under the internal
  // service name, and the file host's check must not cost it the chart engine.
  const origin = getFileTransferOrigin();
  if (!origin) return null;
  return `${origin}${url.pathname}${url.search}`;
}

/**
 * Fetch a file URL; only the configured file host's own routes get the
 * browser's credentials. When that host refuses the file session itself
 * (lib/file-access-repair.ts: its gate cookie is missing), the session is
 * reconnected from the page and the file asked for once more; if that is
 * not possible the refusal is returned as it came, and the shell has said
 * what to do (FileAccessGuard).
 */
export async function fetchPrivateFile(value: string, init: RequestInit = {}): Promise<Response> {
  const owned = isPrivateFileUrl(value, init.method || "GET");
  const send = () => fetch(value, { ...init, credentials: owned ? "include" : "omit", redirect: "error" });
  const response = await send();
  if (!owned || response.status !== 401) return response;
  const recovery = await recoverFromFileGateRefusal(response);
  if (recovery !== "retry" || !requestBodyIsReplayable(init)) return response;
  const again = await send();
  if (again.status === 401 && (await readFileGateRefusal(again)) === "access") noteFileAccessRequired();
  return again;
}
