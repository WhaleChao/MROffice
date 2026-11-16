"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Theme store — persists the user's chosen color scheme to localStorage,
 * sets `data-theme-mode` on `document.documentElement`, and provides a
 * React hook via `useSyncExternalStore` (SSR-safe, no provider needed).
 *
 * Token convention:
 *   data-theme="console"          — always present in the workspace layout
 *   data-theme-mode="dark"        — added when effective theme is dark
 *   data-theme-mode="high-contrast" — added when effective theme is high-contrast
 *   (attribute absent)            — light (default)
 *
 * The `resolveTheme` function is intentionally pure and exported for unit
 * tests; it contains all the precedence logic.
 */

import { useSyncExternalStore } from "react";

import { THEME_KEY } from "./theme-init";

export type ThemeChoice = "light" | "dark" | "high-contrast" | "system";
export type EffectiveTheme = "light" | "dark" | "high-contrast";

const THEME_EVENT = "app-theme-change";

/**
 * Resolve the effective (rendered) theme given a stored choice and the
 * OS-level dark-mode preference.  Pure — no side effects.
 *
 * @param stored   - value read from localStorage, or null when absent
 * @param osDark   - true when `prefers-color-scheme: dark` matches
 */
export function resolveTheme(
  stored: string | null,
  osDark: boolean,
): EffectiveTheme {
  if (stored === "high-contrast") return "high-contrast";
  if (stored === "dark") return "dark";
  if (stored === "light") return "light";
  // Only an EXPLICIT "system" choice follows the OS. No stored preference
  
  
  if (stored === "system") return osDark ? "dark" : "light";
  return "light";
}

// ── Internal helpers ──────────────────────────────────────────────────────────

function readOsDark(): boolean {
  if (typeof window === "undefined") return false;
  if (typeof window.matchMedia !== "function") return false;
  return window.matchMedia("(prefers-color-scheme: dark)").matches;
}

function readStored(): string | null {
  if (typeof window === "undefined") return null;
  return localStorage.getItem(THEME_KEY);
}

/** Light-only surfaces mounted right now (the marketing site). */
let lightHolds = 0;

function computeEffective(): EffectiveTheme {
  if (lightHolds > 0) return "light";
  return resolveTheme(readStored(), readOsDark());
}

/** Applies `data-theme-mode` to <html> and fires the custom event. */
function applyTheme(effective: EffectiveTheme): void {
  if (typeof document === "undefined") return;
  if (effective === "light") {
    document.documentElement.removeAttribute("data-theme-mode");
  } else {
    document.documentElement.setAttribute("data-theme-mode", effective);
  }
  // Also sync the legacy `.dark` class for window-manager.css compatibility
  if (effective === "dark" || effective === "high-contrast") {
    document.documentElement.classList.add("dark");
  } else {
    document.documentElement.classList.remove("dark");
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

/** Puts the current theme back on <html>: a locale switch remounts the root
 * layout, and React clears <html>'s attributes and classes as it does
 * (components/ThemeBootScript). */
export function reapplyTheme(): void {
  applyTheme(computeEffective());
}

/**
 * Persist the user's choice and immediately apply the resulting theme.
 * Pass "system" to store an explicit follow-the-OS choice.
 */
export function setThemeChoice(choice: ThemeChoice): void {
  if (typeof window === "undefined") return;
  // Keep the explicit `system` sentinel. Absence means the owner-defined
  // cream/light default, while an explicit system choice follows the OS;
  // deleting the key here made those two user intents indistinguishable.
  localStorage.setItem(THEME_KEY, choice);
  applyTheme(computeEffective());
  window.dispatchEvent(new Event(THEME_EVENT));
}

export function getThemeChoice(): ThemeChoice {
  const stored = readStored();
  if (
    stored === "dark" ||
    stored === "light" ||
    stored === "high-contrast" ||
    stored === "system"
  ) {
    return stored;
  }
  // Nothing chosen renders light (resolveTheme), so say so: reporting
  // "system" showed a dark-OS user "System" over a light page.
  return "light";
}

/**
 * Hold the light theme while a light-only surface is mounted (the marketing
 * pages have no dark design): <html> loses data-theme-mode and every
 * useEffectiveTheme / useIsDark reads light, so an editor embedded in the
 * page renders light too. Returns the release; the user's own theme comes
 * back when the last hold lets go.
 */
export function holdLightTheme(): () => void {
  if (typeof window === "undefined") return () => {};
  lightHolds += 1;
  applyTheme("light");
  window.dispatchEvent(new Event(THEME_EVENT));
  let released = false;
  return () => {
    if (released) return;
    released = true;
    lightHolds = Math.max(0, lightHolds - 1);
    applyTheme(computeEffective());
    window.dispatchEvent(new Event(THEME_EVENT));
  };
}

// ── useSyncExternalStore plumbing ─────────────────────────────────────────────

function subscribe(onStoreChange: () => void): () => void {
  if (typeof window === "undefined") return () => {};
  const handleStorage = (e: Event) => {
    if (e instanceof StorageEvent && e.key !== THEME_KEY) return;
    onStoreChange();
  };
  const handleColorScheme = () => onStoreChange();
  window.addEventListener("storage", handleStorage);
  window.addEventListener(THEME_EVENT, onStoreChange);
  let mql: MediaQueryList | null = null;
  if (typeof window.matchMedia === "function") {
    mql = window.matchMedia("(prefers-color-scheme: dark)");
    mql.addEventListener("change", handleColorScheme);
  }
  return () => {
    window.removeEventListener("storage", handleStorage);
    window.removeEventListener(THEME_EVENT, onStoreChange);
    mql?.removeEventListener("change", handleColorScheme);
  };
}

/**
 * Returns the current effective theme (light | dark | high-contrast).
 * Safe to call during SSR — returns "light" server-side.
 */
export function useEffectiveTheme(): EffectiveTheme {
  return useSyncExternalStore(subscribe, computeEffective, () => "light");
}

/** Convenience hook — true when the effective theme is dark or high-contrast. */
/**
 * The stored choice as a hook: the server and the hydrating client both
 * render "light", then the client's real choice — reading localStorage in
 * render made the settings page fail hydration (server "light"/"system",
 * client "dark") and leave the select on the server's value.
 */
export function useThemeChoice(): ThemeChoice {
  return useSyncExternalStore(subscribe, getThemeChoice, () => "light");
}

export function useIsDark(): boolean {
  const theme = useEffectiveTheme();
  return theme === "dark" || theme === "high-contrast";
}

// themeInitScript() + THEME_KEY now live in ./theme-init (server-safe, no
// "use client") so the server layout can emit the pre-hydration <script>
// without invoking a client-boundary function.
