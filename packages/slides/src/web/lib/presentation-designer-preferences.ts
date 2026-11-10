/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint Designer's two options (File > Options > General > PowerPoint
 * Designer), for every deck in this browser, kept the way the editors keep
 * AutoSave (lib/office-autosave-preference):
 *  - "Automatically show me design ideas": the pane opens by itself after a
 *    picture, chart, table or list goes on a slide;
 *  - "Automatically show me suggestions when I create a new presentation":
 *    the pane opens once the title slide of a new deck has its title.
 * Both are on by default, as in PowerPoint; turned off, Design > Designer
 * still opens the pane.
 */

export type PresentationDesignerPreference = "autoShow" | "autoShowNewDeck";

const STORAGE_KEYS: Readonly<Record<PresentationDesignerPreference, string>> = {
  autoShow: "mrai.presentation.designer.auto-show",
  autoShowNewDeck: "mrai.presentation.designer.auto-show-new-deck",
};

const cache = new Map<PresentationDesignerPreference, boolean>();
const listeners = new Set<() => void>();

function read(key: PresentationDesignerPreference): boolean {
  if (typeof window === "undefined") return true;
  try {
    return window.localStorage.getItem(STORAGE_KEYS[key]) !== "off";
  } catch {
    return true;
  }
}

export function presentationDesignerPreference(key: PresentationDesignerPreference): boolean {
  let value = cache.get(key);
  if (value === undefined) {
    value = read(key);
    cache.set(key, value);
  }
  return value;
}

export function setPresentationDesignerPreference(key: PresentationDesignerPreference, enabled: boolean): void {
  cache.set(key, enabled);
  try {
    if (enabled) window.localStorage.removeItem(STORAGE_KEYS[key]);
    else window.localStorage.setItem(STORAGE_KEYS[key], "off");
  } catch {
    // Kept for this session only.
  }
  for (const listener of [...listeners]) listener();
}

/** Both options as one snapshot (stable while neither changes). */
let snapshot: Readonly<Record<PresentationDesignerPreference, boolean>> | null = null;
export function presentationDesignerPreferences(): Readonly<Record<PresentationDesignerPreference, boolean>> {
  const autoShow = presentationDesignerPreference("autoShow");
  const autoShowNewDeck = presentationDesignerPreference("autoShowNewDeck");
  if (!snapshot || snapshot.autoShow !== autoShow || snapshot.autoShowNewDeck !== autoShowNewDeck) {
    snapshot = { autoShow, autoShowNewDeck };
  }
  return snapshot;
}

const SERVER_SNAPSHOT: Readonly<Record<PresentationDesignerPreference, boolean>> = { autoShow: true, autoShowNewDeck: true };
export function presentationDesignerPreferencesServerSnapshot(): Readonly<Record<PresentationDesignerPreference, boolean>> {
  return SERVER_SNAPSHOT;
}

export function subscribePresentationDesignerPreferences(listener: () => void): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    const key = (Object.keys(STORAGE_KEYS) as PresentationDesignerPreference[]).find((name) => STORAGE_KEYS[name] === event.key);
    if (!key) return;
    cache.set(key, read(key));
    listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}
