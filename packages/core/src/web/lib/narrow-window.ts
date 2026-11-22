"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";








export const PHONE_DEVICE_QUERY =
  "screen and (pointer: coarse) and (max-device-width: 743px), screen and (pointer: coarse) and (max-device-height: 743px)";









export const NARROW_WINDOW_QUERY = `(max-width: 743px), ${PHONE_DEVICE_QUERY}`;

function subscribe(onChange: () => void): () => void {
  const media = window.matchMedia?.(NARROW_WINDOW_QUERY);
  media?.addEventListener?.("change", onChange);
  return () => media?.removeEventListener?.("change", onChange);
}

function read(): boolean {
  return Boolean(window.matchMedia?.(NARROW_WINDOW_QUERY).matches);
}

/** Live: is this window narrow? Server render and first paint: no. */
export function useNarrowWindow(): boolean {
  return useSyncExternalStore(subscribe, read, () => false);
}

function subscribePhone(onChange: () => void): () => void {
  const media = window.matchMedia?.(PHONE_DEVICE_QUERY);
  media?.addEventListener?.("change", onChange);
  return () => media?.removeEventListener?.("change", onChange);
}

function readPhone(): boolean {
  return Boolean(window.matchMedia?.(PHONE_DEVICE_QUERY).matches);
}

/** Live: is this device a phone? Server render and first paint: no. */
export function usePhoneDevice(): boolean {
  return useSyncExternalStore(subscribePhone, readPhone, () => false);
}

/**
 * A tablet held upright: wider than a phone, narrower than a landscape iPad
 * (a 13-inch iPad upright is 1032 wide). Under a finger only (the caller adds
 * Touch mode); a narrow desktop window is not a tablet.
 */
export const PORTRAIT_TABLET_QUERY =
  "(orientation: portrait) and (min-width: 744px) and (max-width: 1100px)";

function subscribePortraitTablet(onChange: () => void): () => void {
  const media = window.matchMedia?.(PORTRAIT_TABLET_QUERY);
  media?.addEventListener?.("change", onChange);
  return () => media?.removeEventListener?.("change", onChange);
}

function readPortraitTablet(): boolean {
  return Boolean(window.matchMedia?.(PORTRAIT_TABLET_QUERY).matches);
}

export function usePortraitTabletWidth(): boolean {
  return useSyncExternalStore(subscribePortraitTablet, readPortraitTablet, () => false);
}







export const PHONE_WIDTH_QUERY = "(max-width: 600px)";

function subscribePhoneWidth(onChange: () => void): () => void {
  const media = window.matchMedia?.(PHONE_WIDTH_QUERY);
  media?.addEventListener?.("change", onChange);
  return () => media?.removeEventListener?.("change", onChange);
}

function readPhoneWidth(): boolean {
  return Boolean(window.matchMedia?.(PHONE_WIDTH_QUERY).matches);
}

export function usePhoneWidth(): boolean {
  return useSyncExternalStore(subscribePhoneWidth, readPhoneWidth, () => false);
}

/**
 * The editors' compact chrome: a phone (any orientation) or a phone-width
 * window, and an embedded editor whose integration asked for the mobile
 * editor (ONLYOFFICE's `type: "mobile"`, the SDK's `ui.layout: "mobile"`) at
 * any width. There the ribbon is folded to its tab row and a tab shows its
 * commands over the page until the page is touched (Office's "Show Tabs
 * Only"), and the title row keeps Undo, Redo and Save with the rest behind
 * "More" -- the room a phone's page needs (ONLYOFFICE's mobile editors keep
 * one row of commands at the top; Word for iPhone, its ribbon in a sheet).
 */
export const COMPACT_CHROME_QUERY = `${PHONE_WIDTH_QUERY}, ${PHONE_DEVICE_QUERY}`;

/** On <html> while an embedding page chose the editor: "on" the mobile
 * editor at any width, "off" the full editor on a phone too (ONLYOFFICE's
 * Switch to Desktop). Absent, the device decides. */
export const MOBILE_LAYOUT_ATTRIBUTE = "data-mobile-layout";

const layoutListeners = new Set<() => void>();

/** The embedded editor's integration asked for the mobile editor (true),
 * the full one (false), or leaves it to the device (null): every
 * compact-chrome reader follows. */
export function setMobileLayoutRequested(choice: boolean | null): void {
  if (typeof document === "undefined") return;
  const root = document.documentElement;
  const next = choice === null ? null : choice ? "on" : "off";
  if (root.getAttribute(MOBILE_LAYOUT_ATTRIBUTE) === next) return;
  if (next === null) root.removeAttribute(MOBILE_LAYOUT_ATTRIBUTE);
  else root.setAttribute(MOBILE_LAYOUT_ATTRIBUTE, next);
  for (const listener of layoutListeners) listener();
}

function subscribeCompactChrome(onChange: () => void): () => void {
  const media = window.matchMedia?.(COMPACT_CHROME_QUERY);
  media?.addEventListener?.("change", onChange);
  layoutListeners.add(onChange);
  return () => {
    media?.removeEventListener?.("change", onChange);
    layoutListeners.delete(onChange);
  };
}

/** Now: does this window get the editors' compact chrome? */
export function compactEditorChrome(): boolean {
  if (typeof window === "undefined") return false;
  const chosen = document.documentElement.getAttribute(MOBILE_LAYOUT_ATTRIBUTE);
  if (chosen === "on") return true;
  if (chosen === "off") return false;
  return Boolean(window.matchMedia?.(COMPACT_CHROME_QUERY).matches);
}

/** Live: does this window get the editors' compact chrome? Server render and
 * first paint: no. */
export function useCompactEditorChrome(): boolean {
  return useSyncExternalStore(subscribeCompactChrome, compactEditorChrome, () => false);
}
