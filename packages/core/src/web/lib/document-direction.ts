/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";

/**
 * The interface's writing direction, as `<html dir>` states it
 * (app/[locale]/layout.tsx sets it from i18n/locales.ts `localeDirection`).
 *
 * Only "rtl" is reported; an ltr page (and the server, and a test document
 * without `dir`) reads `false` / `undefined`, so code that branches on it
 * keeps its ltr path exactly as it was. Document content has its own
 * direction and is not asked here.
 */
export function isDocumentRtl(): boolean {
  return typeof document !== "undefined" && document.documentElement.dir === "rtl";
}

const listeners = new Set<() => void>();
let observer: MutationObserver | null = null;

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  // One observer for every subscriber: a ribbon mounts hundreds of menus.
  if (!observer && typeof MutationObserver !== "undefined" && typeof document !== "undefined") {
    observer = new MutationObserver(() => {
      for (const notify of listeners) notify();
    });
    observer.observe(document.documentElement, { attributes: true, attributeFilter: ["dir"] });
  }
  return () => {
    listeners.delete(listener);
    if (listeners.size === 0 && observer) {
      observer.disconnect();
      observer = null;
    }
  };
}

const clientDirection = (): "rtl" | undefined => (isDocumentRtl() ? "rtl" : undefined);
const serverDirection = (): undefined => undefined;

/**
 * `"rtl"` under an rtl interface, else `undefined` (a component's own ltr
 * default). Hydrates as `undefined` and settles on the client, so the server
 * markup never disagrees with the first client render.
 */
export function useDocumentDirection(): "rtl" | undefined {
  return useSyncExternalStore(subscribe, clientDirection, serverDirection);
}
