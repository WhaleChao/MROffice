/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Server-safe theme bootstrap. There is intentionally NO "use client" here:
 * the server layout imports {@link themeInitScript} to emit a pre-hydration
 * <script>, so this module must not pull in React or any client-only code.
 * The client store (`theme.ts`) re-imports {@link THEME_KEY} from here so both
 * sides agree on the localStorage key.
 */

import { productSurfaceInitStatements } from "./product-surface";

export const THEME_KEY = "app-theme-choice";

/**
 * CSP source hash of {@link themeInitScript}. The root layout is rendered
 * statically, so it cannot read the per-request nonce the production CSP
 * (proxy.ts) issues; without this hash the bootstrap was refused
 * ("Executing inline script violates ... script-src 'self' 'nonce-…'
 * 'strict-dynamic'") and the stored theme was not applied before paint.
 * tests/unit/proxy-auth.test.ts recomputes it from the script.
 */
export const THEME_INIT_SCRIPT_SHA256 = "kzpd61f2eI34vR8oQZ4Vl+bEZlfzk0ZFH2kwQhxTRq0=";

/**
 * Inline IIFE string injected before React hydration so there is no flash of
 * the wrong theme. Reads the stored choice (or the OS preference) and sets
 * `data-theme-mode` + the legacy `.dark` class on <html>. It also marks the
 * platform on <html> (`data-platform`: "apple" for macOS, iOS and iPadOS,
 * "other" for Windows, Linux, Android), before the first paint, so a style
 * can answer the platform's scrollbars -- Windows draws classic, always-visible
 * bars where Apple's overlay -- without a flash of the browser's own. Last,
 * on MR Office's own host it marks the product surface (`data-surface`,
 * lib/product-surface), in its own try so a refused storage read changes
 * nothing above. Pure string output.
 */
export function themeInitScript(): string {
  return `(function(){try{var p=(navigator.userAgentData&&navigator.userAgentData.platform)||navigator.platform||navigator.userAgent||"";document.documentElement.setAttribute("data-platform",/mac|iphone|ipad|ipod/i.test(p)?"apple":"other");var s=localStorage.getItem(${JSON.stringify(THEME_KEY)});var osDark=window.matchMedia&&window.matchMedia("(prefers-color-scheme: dark)").matches;var e=(s==="high-contrast")?"high-contrast":(s==="dark")?"dark":(s==="light")?"light":(s==="system"&&osDark)?"dark":"light";if(e!=="light"){document.documentElement.setAttribute("data-theme-mode",e);document.documentElement.classList.add("dark");}else{document.documentElement.removeAttribute("data-theme-mode");document.documentElement.classList.remove("dark");}}catch(err){}try{${productSurfaceInitStatements()}}catch(err){}})();`;
}
