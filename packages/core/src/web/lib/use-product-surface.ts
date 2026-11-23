"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";

import { productSurfaceOfHost, SURFACE_OVERRIDE_KEY, type ProductSurface } from "./product-surface";

const subscribe = () => () => {};

const LOOPBACK_HOSTS = ["localhost", "127.0.0.1", "[::1]"];

/** The surface as the bootstrap stamped it (lib/product-surface), or as the
 * host says by the same rule: a locale switch strips <html> of the stamp
 * until components/ThemeBootScript puts it back, and what renders between
 * must not read MR Note on MR Office's host. */
export function readProductSurface(): ProductSurface {
  if (typeof document === "undefined") return "note";
  if (document.documentElement.getAttribute("data-surface") === "office") return "office";
  const host = location.hostname;
  if (productSurfaceOfHost(host) === "office") return "office";
  if (!LOOPBACK_HOSTS.includes(host)) return "note";
  try {
    return localStorage.getItem(SURFACE_OVERRIDE_KEY) === "office" ? "office" : "note";
  } catch {
    return "note";
  }
}

/**
 * The product surface of this visit. The server render and the hydrating
 * render are MR Note's (no server render knows the host); a visit on the
 * office host re-renders as MR Office right after hydration. The notebook
 * interior, where the surfaces differ, renders on the client only, so it is
 * never drawn as MR Note first.
 */
export function useProductSurface(): ProductSurface {
  return useSyncExternalStore<ProductSurface>(subscribe, readProductSurface, () => "note");
}
