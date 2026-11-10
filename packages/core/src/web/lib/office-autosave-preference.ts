/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const STORAGE_KEY = "mrai.office.autosave";

const cached = new Map<string, boolean>();
const listeners = new Set<() => void>();

function storageKey(scope?: string): string {
  return scope ? `${STORAGE_KEY}.${scope}` : STORAGE_KEY;
}

function read(key: string): boolean {
  if (typeof window === "undefined") return true;
  try {
    return window.localStorage.getItem(key) !== "off";
  } catch {
    return true;
  }
}

export function officeAutosaveEnabled(scope?: string): boolean {
  const key = storageKey(scope);
  let value = cached.get(key);
  if (value === undefined) {
    value = read(key);
    cached.set(key, value);
  }
  return value;
}

export function setOfficeAutosaveEnabled(enabled: boolean, scope?: string): void {
  const key = storageKey(scope);
  cached.set(key, enabled);
  try {
    if (enabled) window.localStorage.removeItem(key);
    else window.localStorage.setItem(key, "off");
  } catch {
    // Kept for this session only.
  }
  for (const listener of [...listeners]) listener();
}

export function subscribeOfficeAutosave(listener: () => void): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key !== STORAGE_KEY && !event.key?.startsWith(`${STORAGE_KEY}.`)) return;
    cached.delete(event.key);
    listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}

/** The four strings the save face shows for the switch. */
export interface OfficeAutosaveLabels {
  readonly on: string;
  readonly off: string;
  readonly holdToTurnOff: string;
  readonly holdToTurnOn: string;
}

/** The labels from a `console-notebooks` translator. */
export function officeAutosaveLabels(t: unknown): OfficeAutosaveLabels {
  const translate = t as (key: string) => string;
  return {
    on: translate("pages.autosaveOn"),
    off: translate("pages.autosaveOff"),
    holdToTurnOff: translate("pages.autosaveHoldToTurnOff"),
    holdToTurnOn: translate("pages.autosaveHoldToTurnOn"),
  };
}
