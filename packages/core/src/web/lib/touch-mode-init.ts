/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Server-safe Touch mode bootstrap (no "use client", no React): the root
 * layout emits {@link touchModeInitScript} before hydration so the first
 * paint already has the reader's Touch mode (lib/touch-mode). Touch sizing
 * follows `data-touch-mode` alone in the workspace (styles/touch-sizing.css),
 * so without this an iPad with the switch off would paint the finger sizes
 * until the workspace mounted, and one with it on the mouse's sizes.
 *
 * It only acts on the workspace's pages, which TouchModeSync then keeps on
 * the switch: it marks <html> `data-workspace-app` (the marker
 * WorkspaceClientLayout keeps there while the workspace is open, and which
 * hands the public pages' touch sizing to the device) and applies the
 * stored choice, or the device's default (lib/touch-mode deviceHasTouch).
 */

export const TOUCH_MODE_KEY = "mrnote.touchMode";
export const TOUCH_MODE_ATTRIBUTE = "data-touch-mode";
export const WORKSPACE_APP_ATTRIBUTE = "data-workspace-app";

/** A workspace page, with or without its locale prefix (`as-needed`). */
export const WORKSPACE_PATH_PATTERN = /^\/(?:[a-z]{2,3}(?:-[A-Za-z0-9]{2,4})?\/)?workspace(?:\/|$)/;

/**
 * CSP source hash of {@link touchModeInitScript}, allowed beside the theme
 * bootstrap's (proxy.ts) because the statically rendered root layout cannot
 * carry the per-request nonce. tests/unit/proxy-auth.test.ts recomputes it.
 */
export const TOUCH_MODE_INIT_SCRIPT_SHA256 = "Y33OJd7fkHk38rTe8k6QlVeKah0ZZWVeqC0yztxx7Ic=";

export function touchModeInitScript(): string {
  return `(function(){try{if(!${WORKSPACE_PATH_PATTERN.toString()}.test(location.pathname))return;var d=document.documentElement;d.setAttribute(${JSON.stringify(WORKSPACE_APP_ATTRIBUTE)},"");var s=null;try{s=localStorage.getItem(${JSON.stringify(TOUCH_MODE_KEY)})}catch(e){}var on=s==="1"?true:s==="0"?false:(navigator.maxTouchPoints||0)>0||!!(window.matchMedia&&window.matchMedia("(any-pointer: coarse)").matches);if(on)d.setAttribute(${JSON.stringify(TOUCH_MODE_ATTRIBUTE)},"");else d.removeAttribute(${JSON.stringify(TOUCH_MODE_ATTRIBUTE)})}catch(e){}})();`;
}
