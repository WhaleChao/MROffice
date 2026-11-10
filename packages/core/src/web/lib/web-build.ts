"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isAccessGateRefusal, noteAccessGateRequired } from "./access-gate";
import { WEB_BUILD_ID } from "./web-build-id.generated";
import { WEB_CLIENT_PROTOCOL } from "./web-protocol";















export const WEB_BUILD_CHECK_EVENT = "mrnote:web-build-check";
/** The page runs an older build; it keeps working and loads the new one at
 * its next page change. */
export const WEB_BUILD_STALE_EVENT = "mrnote:web-build-stale";
/** The page must reload now: the server no longer accepts what it sends (a
 * protocol below the server's minimum, a save refused on an old build, code
 * or a stylesheet that is gone). The notice with its button shows. */
export const WEB_BUILD_RELOAD_NEEDED_EVENT = "mrnote:web-build-reload-needed";
const ENDPOINT = "/web-build.json";
/** A check requested less than this long after the previous one is skipped. */
const MIN_CHECK_INTERVAL_MS = 30_000;

let stale = false;
let reloadNeeded = false;
let lastCheckAt = 0;
let inFlight: Promise<boolean> | null = null;
const listeners = new Set<(stale: boolean) => void>();

export function currentWebBuildId(): string {
  return WEB_BUILD_ID;
}

export function isWebBuildStale(): boolean {
  return stale;
}

export function subscribeWebBuildStale(listener: (stale: boolean) => void): () => void {
  listeners.add(listener);
  return () => { listeners.delete(listener); };
}

/** Stale when both sides know their build and they differ. */
export function webBuildIsStale(page: string, server: unknown): boolean {
  return page !== "" && typeof server === "string" && server !== "" && server !== page;
}

/** The server says this page's protocol is older than it accepts. A server
 * that publishes no minimum (built before it did) accepts every page. */
export function webBuildIsIncompatible(pageProtocol: number, serverMinimum: unknown): boolean {
  return typeof serverMinimum === "number" && Number.isFinite(serverMinimum) && pageProtocol < serverMinimum;
}

/** Refusals that, on a page of an older build, mean it must reload now. */
export function refusalNeedsReload(reason: string): boolean {
  return reason.startsWith("save_") || reason.startsWith("page_causal") || reason === "window_code_unavailable";
}

export function isWebBuildReloadNeeded(): boolean {
  return reloadNeeded;
}

/** Ask the page to reload now (the notice shows once, and stays). */
export function markWebBuildReloadNeeded(): void {
  if (reloadNeeded) return;
  reloadNeeded = true;
  if (typeof window !== "undefined") window.dispatchEvent(new CustomEvent(WEB_BUILD_RELOAD_NEEDED_EVENT));
}

function markStale(): void {
  if (stale) return;
  stale = true;
  for (const listener of listeners) listener(true);
  if (typeof window !== "undefined") window.dispatchEvent(new CustomEvent(WEB_BUILD_STALE_EVENT));
}

/**
 * Compare this page's build with the server's. Resolves to whether the page
 * is stale. Coalesces concurrent calls and, unless `force`, skips a check
 * made shortly after the last one. Network failures resolve to the last
 * verdict: an unreachable server is not evidence of a new build.
 */
export function checkWebBuildFreshness(options: { force?: boolean } = {}): Promise<boolean> {
  // A stale page still asks now and then: a later release may raise the
  // minimum protocol past it. Once it must reload, there is nothing to learn.
  if (reloadNeeded || WEB_BUILD_ID === "" || typeof fetch !== "function") return Promise.resolve(stale);
  if (inFlight) return inFlight;
  const now = Date.now();
  if (!options.force && now - lastCheckAt < MIN_CHECK_INTERVAL_MS) return Promise.resolve(stale);
  lastCheckAt = now;
  // Sent with the site's own cookies: behind the private-beta access gate an
  // anonymous request is answered by the gate (401) and never reaches this
  // route, which left every unlocked page unable to learn of a new build.
  inFlight = fetch(ENDPOINT, { cache: "no-store", credentials: "same-origin", headers: { accept: "application/json" } })
    .then(async (response) => {
      if (!response.ok) {
        // The pass itself has run out: say so instead of staying silent.
        const refusal: unknown = await response.json().catch(() => null);
        if (isAccessGateRefusal(response.status, refusal)) noteAccessGateRequired();
        return stale;
      }
      const body = (await response.json()) as { build?: unknown; minProtocol?: unknown };
      if (webBuildIsStale(WEB_BUILD_ID, body?.build)) {
        markStale();
        if (webBuildIsIncompatible(WEB_CLIENT_PROTOCOL, body?.minProtocol)) markWebBuildReloadNeeded();
      }
      return stale;
    })
    .catch(() => stale)
    .finally(() => { inFlight = null; });
  return inFlight;
}

/**
 * Ask for a freshness check from code that must not import the checker's
 * state (the autosave hook, on a refusal an old build would produce). The
 * console's guard listens and shows the way out when the page is stale.
 */
export function requestWebBuildCheck(reason: string): void {
  if (typeof window === "undefined") return;
  window.dispatchEvent(new CustomEvent(WEB_BUILD_CHECK_EVENT, { detail: { reason } }));
}

/** Test seam: forget the verdict and the rate limit. */
export function resetWebBuildFreshnessForTests(): void {
  stale = false;
  reloadNeeded = false;
  lastCheckAt = 0;
  inFlight = null;
}
