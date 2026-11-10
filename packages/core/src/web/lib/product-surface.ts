/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type ProductSurface = "note" | "office";

/** The hosts MR Office is served on. */
export const OFFICE_SURFACE_HOSTS: readonly string[] = ["mroffice.ai", "www.mroffice.ai"];

/** What a developer's browser keeps the chosen surface under (loopback
 * hosts only): `?surface=office` sets it, `?surface=note` clears it. */
export const SURFACE_OVERRIDE_KEY = "mr-surface";

function bareHost(host: string | null | undefined): string {
  return (host ?? "").trim().toLowerCase().replace(/:\d+$/, "").replace(/\.$/, "");
}

/** The surface a host serves: MR Office on its own hosts, MR Note elsewhere. */
export function productSurfaceOfHost(host: string | null | undefined): ProductSurface {
  return OFFICE_SURFACE_HOSTS.includes(bareHost(host)) ? "office" : "note";
}

/**
 * Statements for the pre-hydration bootstrap: stamp the office surface on
 * <html>. A loopback host (a developer's own machine) may ask for it with
 * `?surface=office`, remembered until `?surface=note`; no public host can.
 */
export function productSurfaceInitStatements(): string {
  const hosts = JSON.stringify(OFFICE_SURFACE_HOSTS);
  const key = JSON.stringify(SURFACE_OVERRIDE_KEY);
  return `var h=location.hostname,o=${hosts}.indexOf(h)>=0;if(!o&&(h==="localhost"||h==="127.0.0.1"||h==="[::1]")){var q=/[?&]surface=(office|note)(&|$)/.exec(location.search);if(q)localStorage.setItem(${key},q[1]);o=localStorage.getItem(${key})==="office";}if(o)document.documentElement.setAttribute("data-surface","office");`;
}
