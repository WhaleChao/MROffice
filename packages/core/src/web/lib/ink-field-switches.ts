/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const KEY = "mrai.inkFieldSwitches";
const NAMES = ["inkdiag", "inkpd", "inkscrib", "inkshield"] as const;

function stored(): Record<string, string> {
  try {
    const raw = window.sessionStorage.getItem(KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : null;
    return parsed && typeof parsed === "object" ? (parsed as Record<string, string>) : {};
  } catch {
    return {};
  }
}

/** Keep the switches the current address names; call as early as possible. */
export function captureInkFieldSwitches(): void {
  if (typeof window === "undefined") return;
  try {
    const query = new URLSearchParams(window.location.search);
    const kept = stored();
    let changed = false;
    for (const name of NAMES) {
      const value = query.get(name);
      if (value === null || kept[name] === value) continue;
      kept[name] = value.slice(0, 16);
      changed = true;
    }
    if (changed) window.sessionStorage.setItem(KEY, JSON.stringify(kept));
  } catch {
    // Best-effort: without storage the address alone is read.
  }
}

/** The switches: the current address over what the tab kept. */
export function inkFieldSwitchQuery(): URLSearchParams {
  const query = new URLSearchParams();
  if (typeof window === "undefined") return query;
  captureInkFieldSwitches();
  const kept = stored();
  for (const name of NAMES) if (typeof kept[name] === "string") query.set(name, kept[name]);
  try {
    const current = new URLSearchParams(window.location.search);
    for (const name of NAMES) {
      const value = current.get(name);
      if (value !== null) query.set(name, value);
    }
  } catch {
    // The kept switches stand.
  }
  return query;
}

captureInkFieldSwitches();
