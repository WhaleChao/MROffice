/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const IS_DESKTOP_CLASSIC = process.env.NEXT_PUBLIC_MRAI_DESKTOP === "classic";

/**
 * Where the app's Python runs (lib/code-runner/python-runner.ts): Pyodide in
 * a frame of an origin other than the app's. On the site that is a sandbox
 * domain of its own (NEXT_PUBLIC_SANDBOX_ORIGIN); in the app it is the same
 * local web server addressed as `localhost` instead of `127.0.0.1`: another
 * host, so another site, which shares no cookie, storage or DOM with the
 * app's pages. That host serves only the relay page, its worker and the
 * Pyodide the app ships (desktop-sandbox.mjs), so Python needs no network.
 *
 * The port is chosen at each launch, so the origin is derived from the
 * page's own host (or, on the server, the request's): `127.0.0.1:<port>`
 * gives `http://localhost:<port>`, anything else gives null.
 */
export function desktopPythonSandboxOrigin(appHost: string | null | undefined): string | null {
  const match = /^127\.0\.0\.1:(\d{1,5})$/.exec(appHost ?? "");
  if (!match) return null;
  const port = Number(match[1]);
  return port > 0 && port < 65536 ? `http://localhost:${port}` : null;
}

/**
 * Where the sandbox host serves the Pyodide distribution the app ships
 * (apps/desktop/scripts/classic-pyodide.mjs stages it at this path; the
 * version is the one python-runner.ts and python-worker.js pin).
 */
export const DESKTOP_PYODIDE_INDEX_PATH = "/pyodide/v0.26.2/full/";
