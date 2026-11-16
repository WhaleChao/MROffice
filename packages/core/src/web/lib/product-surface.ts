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










export function isNoteOnlyPath(path: string): boolean {
  const bare = path.length > 1 ? path.replace(/\/+$/, "") : path;
  return (
    bare === "/" ||
    /^\/(?:app|workspace)(?:\/notebooks)?$/.test(bare) ||
    /^\/(?:app|workspace)\/(?:messages|monitor)(?:\/|$)/.test(bare) ||
    /^\/(?:app|workspace)\/notebooks\/[^/]+\/(?:chat|memory)(?:\/|$)/.test(bare)
  );
}

/** The legal documents both products share, at MR Note's addresses; MR
 * Office shows them at /office/legal/<slug>, without MR Note's frame. */
export const OFFICE_LEGAL_PATHS = {
  terms: "/terms",
  privacy: "/privacy",
  cookies: "/cookies",
  "acceptable-use": "/acceptable-use",
  security: "/security/policy",
} as const;

export type OfficeLegalSlug = keyof typeof OFFICE_LEGAL_PATHS;

/** MR Note's own marketing pages, about what MR Office does not have. */
const NOTE_MARKETING_PATH =
  /^\/(?:calendar|memory|workbench|changelog|security|accessibility)(?:\/|$)/;












export function officeRedirectOf(path: string, signedIn: boolean): string | null {
  const bare = path.length > 1 ? path.replace(/\/+$/, "") : path;
  if (isNoteOnlyPath(bare)) return signedIn ? "/app/recent" : "/app/notebooks/guest";
  if (bare === "/pricing") return signedIn ? "/app/settings#billing" : "/office/about";
  for (const [slug, legalPath] of Object.entries(OFFICE_LEGAL_PATHS)) {
    if (bare === legalPath) return `/office/legal/${slug}`;
  }
  if (bare === "/download" || bare === "/about") return `/office${bare}`;
  if (NOTE_MARKETING_PATH.test(bare)) return "/office/about";
  return null;
}

/** The surface a host serves: MR Office on its own hosts, MR Note elsewhere. */
export function productSurfaceOfHost(host: string | null | undefined): ProductSurface {
  return OFFICE_SURFACE_HOSTS.includes(bareHost(host)) ? "office" : "note";
}












export function productSurfaceInitStatements(): string {
  const hosts = JSON.stringify(OFFICE_SURFACE_HOSTS);
  const key = JSON.stringify(SURFACE_OVERRIDE_KEY);
  return `var h=location.hostname,o=${hosts}.indexOf(h)>=0;if(!o&&(h==="localhost"||h==="127.0.0.1"||h==="[::1]")){var q=/[?&]surface=(office|note)(&|$)/.exec(location.search);if(q)localStorage.setItem(${key},q[1]);o=localStorage.getItem(${key})==="office";}if(o){document.documentElement.setAttribute("data-surface","office");var r=function(){var t=document.title,n=t.replace(/MR ?Note/g,"MR Office");if(n!==t)document.title=n;};r();new MutationObserver(r).observe(document.head,{childList:true,subtree:true,characterData:true});}`;
}
