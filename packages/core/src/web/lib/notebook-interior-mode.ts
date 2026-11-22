"use client";























/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type NotebookInteriorMode = "classic" | "workbench";

/** What the mode is decided from; null while it is not known yet. */
export interface NotebookInteriorFacts {
  /** The notebook is in the person's own workspace. */
  readonly own: boolean | null;
  /** The person's effective plan code ("free", "plus", "pro", ...). */
  readonly plan: string | null;
  /** The person's saved choice (this device's, else the account's). */
  readonly preference: NotebookInteriorMode | null;
}

export interface NotebookInteriorResolution {
  /** null until the facts it needs are known. */
  readonly mode: NotebookInteriorMode | null;
  /** The person may switch this notebook between the two designs. */
  readonly canSwitch: boolean;
}

export function isPaidPlan(plan: string | null | undefined): boolean {
  return Boolean(plan) && plan !== "free";
}

/** The pure rule; see the module comment. */
export function resolveNotebookInteriorMode(
  facts: NotebookInteriorFacts,
): NotebookInteriorResolution {
  if (facts.own === false) {
    return { mode: facts.preference ?? "workbench", canSwitch: true };
  }
  if (facts.own === null || facts.plan === null) {
    return { mode: null, canSwitch: false };
  }
  return {
    mode: facts.preference ?? (isPaidPlan(facts.plan) ? "workbench" : "classic"),
    canSwitch: true,
  };
}




export function resolveOfficeInteriorMode(
  choice: NotebookInteriorMode | null,
): NotebookInteriorResolution {
  return { mode: choice ?? "classic", canSwitch: true };
}

// ---------------------------------------------------------------------------
// Stored values: the person's preference and the last mode each notebook
// showed (so a reload paints the right design before the plan is read).
// ---------------------------------------------------------------------------

const PREFERENCE_KEY_PREFIX = "mrnote.notebookInterior.preference.";
const DEVICE_KEY_PREFIX = "mrnote.notebookInterior.device.";
const LAST_MODE_KEY_PREFIX = "mrnote.notebookInterior.last.";

export const NOTEBOOK_INTERIOR_PREFERENCE_EVENT =
  "mrnote:notebook-interior-preference";

function asMode(value: unknown): NotebookInteriorMode | null {
  return value === "classic" || value === "workbench" ? value : null;
}

function readStorage(key: string): string | null {
  if (typeof window === "undefined") return null;
  try {
    return window.localStorage.getItem(key);
  } catch {
    return null;
  }
}

function writeStorage(key: string, value: string | null): void {
  if (typeof window === "undefined") return;
  try {
    if (value === null) window.localStorage.removeItem(key);
    else window.localStorage.setItem(key, value);
  } catch {
    /* no storage: the value lasts as long as the page */
  }
}

/** The person's choice for their own notebooks, per account. */
export function readNotebookInteriorPreference(
  accountKey: string,
): NotebookInteriorMode | null {
  return asMode(readStorage(PREFERENCE_KEY_PREFIX + accountKey));
}

/** A changed choice makes every notebook's remembered design stale: the next
 * visit paints from the choice, not from what it showed before. */
function forgetLastNotebookInteriorModes(): void {
  if (typeof window === "undefined") return;
  try {
    const stale: string[] = [];
    for (let index = 0; index < window.localStorage.length; index += 1) {
      const key = window.localStorage.key(index);
      if (key?.startsWith(LAST_MODE_KEY_PREFIX)) stale.push(key);
    }
    for (const key of stale) window.localStorage.removeItem(key);
  } catch {
    /* no storage */
  }
}

/** The key of this browser's copy of an account's choice (other tabs hear
 * its changes through the storage event). */
export function isNotebookInteriorPreferenceKey(key: string | null): boolean {
  return Boolean(key?.startsWith(PREFERENCE_KEY_PREFIX));
}

export function writeNotebookInteriorPreference(
  accountKey: string,
  mode: NotebookInteriorMode,
): void {
  if (readNotebookInteriorPreference(accountKey) !== mode) forgetLastNotebookInteriorModes();
  writeStorage(PREFERENCE_KEY_PREFIX + accountKey, mode);
  if (typeof window !== "undefined") {
    window.dispatchEvent(
      new CustomEvent(NOTEBOOK_INTERIOR_PREFERENCE_EVENT, {
        detail: { accountKey, mode },
      }),
    );
  }
}

/**
 * The design switched to inside a notebook on THIS device. It is the
 * device's own: the assistant in classic tabs on the computer and the
 * workbench on the iPad, at once, is one person's ordinary day. It stands
 * over the account's choice here; the account's choice (Settings) stays the
 * default of a device that never switched.
 */
export function readDeviceNotebookInterior(
  accountKey: string,
): NotebookInteriorMode | null {
  return asMode(readStorage(DEVICE_KEY_PREFIX + accountKey));
}

export function writeDeviceNotebookInterior(
  accountKey: string,
  mode: NotebookInteriorMode | null,
): void {
  if (readDeviceNotebookInterior(accountKey) === mode) return;
  forgetLastNotebookInteriorModes();
  writeStorage(DEVICE_KEY_PREFIX + accountKey, mode);
  if (typeof window !== "undefined") {
    window.dispatchEvent(
      new CustomEvent(NOTEBOOK_INTERIOR_PREFERENCE_EVENT, {
        detail: { accountKey, mode, device: true },
      }),
    );
  }
}

/** The key of this device's own switch (other tabs hear it change). */
export function isDeviceNotebookInteriorKey(key: string | null): boolean {
  return Boolean(key?.startsWith(DEVICE_KEY_PREFIX));
}

/** Forget this browser's copy (the account says nothing is chosen). */
export function clearNotebookInteriorPreference(accountKey: string): void {
  if (readNotebookInteriorPreference(accountKey) === null) return;
  forgetLastNotebookInteriorModes();
  writeStorage(PREFERENCE_KEY_PREFIX + accountKey, null);
  if (typeof window !== "undefined") {
    window.dispatchEvent(
      new CustomEvent(NOTEBOOK_INTERIOR_PREFERENCE_EVENT, {
        detail: { accountKey, mode: null },
      }),
    );
  }
}

const OFFICE_CHOICE_KEY_PREFIX = "mroffice.workspaceInterior.";

/** The design an MR Office workspace was last switched to on this device. */
export function readOfficeNotebookInterior(
  notebookId: string,
): NotebookInteriorMode | null {
  return asMode(readStorage(OFFICE_CHOICE_KEY_PREFIX + notebookId));
}

export function writeOfficeNotebookInterior(
  notebookId: string,
  mode: NotebookInteriorMode,
): void {
  if (readOfficeNotebookInterior(notebookId) === mode) return;
  writeStorage(OFFICE_CHOICE_KEY_PREFIX + notebookId, mode);
  if (typeof window !== "undefined") {
    window.dispatchEvent(
      new CustomEvent(NOTEBOOK_INTERIOR_PREFERENCE_EVENT, {
        detail: { notebookId, mode, office: true },
      }),
    );
  }
}

/** The key of an MR Office workspace's choice (other tabs hear it change). */
export function isOfficeNotebookInteriorKey(key: string | null): boolean {
  return Boolean(key?.startsWith(OFFICE_CHOICE_KEY_PREFIX));
}

export function readLastNotebookInteriorMode(
  notebookId: string,
): NotebookInteriorMode | null {
  return asMode(readStorage(LAST_MODE_KEY_PREFIX + notebookId));
}

export function writeLastNotebookInteriorMode(
  notebookId: string,
  mode: NotebookInteriorMode,
): void {
  if (readLastNotebookInteriorMode(notebookId) === mode) return;
  writeStorage(LAST_MODE_KEY_PREFIX + notebookId, mode);
}

// ---------------------------------------------------------------------------
// Tabs
// ---------------------------------------------------------------------------

/** The window types that are not tabs in classic mode: the assistant is the
 * notebook's side panel there, beside whichever tab is in front. */
export const CLASSIC_NON_TAB_TYPES: ReadonlySet<string> = new Set(["ai_panel"]);




export const CLASSIC_TAB_META_KEY = "classicTab";

export interface ClassicTabWindow {
  readonly id: string;
  readonly type: string;
  readonly zIndex: number;
  readonly minimized: boolean;
  readonly meta?: Readonly<Record<string, string>>;
}

/** The assistant stands as a tab of its own (CLASSIC_TAB_META_KEY). */
export function assistantIsClassicTab(windowState: ClassicTabWindow): boolean {
  return windowState.meta?.[CLASSIC_TAB_META_KEY] === "1";
}

/**
 * A window that is one of the tab strip's tabs. In a narrow window there is
 * no room beside a page for the assistant, so it is a tab there too.
 */
export function isClassicTabWindow(
  windowState: ClassicTabWindow,
  assistantIsTab = false,
): boolean {
  if (
    !assistantIsTab &&
    CLASSIC_NON_TAB_TYPES.has(windowState.type) &&
    !assistantIsClassicTab(windowState)
  ) {
    return false;
  }
  // A Docs chart's workbook is kept mounted while minimized, as a helper
  // of its document: it is not a tab of its own until it is shown.
  return !(windowState.minimized && windowState.type === "chart_workbook");
}

/**
 * The tabs in their strip order: the stored order first (what the person
 * dragged them into), then any window it does not name, in the desk's own
 * order (the order they were opened).
 */
export function orderClassicTabs<W extends ClassicTabWindow>(
  windows: readonly W[],
  storedOrder: readonly string[],
  assistantIsTab = false,
): W[] {
  const tabs = windows.filter((windowState) => isClassicTabWindow(windowState, assistantIsTab));
  const rank = new Map(storedOrder.map((id, index) => [id, index] as const));
  return tabs
    .map((windowState, index) => ({ windowState, index }))
    .sort((a, b) => {
      const ra = rank.get(a.windowState.id);
      const rb = rank.get(b.windowState.id);
      if (ra !== undefined && rb !== undefined) return ra - rb;
      if (ra !== undefined) return -1;
      if (rb !== undefined) return 1;
      return a.index - b.index;
    })
    .map(({ windowState }) => windowState);
}

/** The tab in front: the topmost tab window that is not minimized. */
export function activeClassicTabId(
  windows: readonly ClassicTabWindow[],
  assistantIsTab = false,
): string | null {
  let top: ClassicTabWindow | null = null;
  for (const windowState of windows) {
    if (!isClassicTabWindow(windowState, assistantIsTab) || windowState.minimized) continue;
    if (!top || windowState.zIndex > top.zIndex) top = windowState;
  }
  return top?.id ?? null;
}

/** Move `id` to `toIndex` of the strip (indices of the ordered tabs). */
export function moveClassicTab(
  orderedIds: readonly string[],
  id: string,
  toIndex: number,
): string[] {
  const from = orderedIds.indexOf(id);
  if (from < 0) return [...orderedIds];
  const next = orderedIds.filter((item) => item !== id);
  const clamped = Math.max(0, Math.min(next.length, toIndex));
  next.splice(clamped, 0, id);
  return next;
}

const TAB_ORDER_KEY_PREFIX = "mrnote.classicTabs.";

export function readClassicTabOrder(layoutKey: string): string[] {
  const raw = readStorage(TAB_ORDER_KEY_PREFIX + layoutKey);
  if (!raw) return [];
  try {
    const parsed: unknown = JSON.parse(raw);
    return Array.isArray(parsed)
      ? parsed.filter((item): item is string => typeof item === "string").slice(0, 500)
      : [];
  } catch {
    return [];
  }
}

export function writeClassicTabOrder(layoutKey: string, order: readonly string[]): void {
  writeStorage(
    TAB_ORDER_KEY_PREFIX + layoutKey,
    order.length > 0 ? JSON.stringify(order.slice(0, 500)) : null,
  );
}

/**
 * The tab a keyboard chord asks for, or null when it is not a tab chord.
 * Ctrl+Tab / Ctrl+Shift+Tab step through the strip (wrapping); Ctrl/Cmd+1..8
 * pick that tab and Ctrl/Cmd+9 the last one, as in a browser. The workbench's
 * own Cmd/Ctrl+Alt+[ and ] are handled by the canvas in strip order too.
 */
export function classicTabForChord(
  event: Pick<KeyboardEvent, "key" | "code" | "ctrlKey" | "metaKey" | "altKey" | "shiftKey">,
  orderedIds: readonly string[],
  activeId: string | null,
): string | null {
  if (orderedIds.length === 0) return null;
  if (event.code === "Tab" || event.key === "Tab") {
    if (!event.ctrlKey || event.metaKey || event.altKey) return null;
    return stepClassicTab(orderedIds, activeId, event.shiftKey ? -1 : 1);
  }
  if (!(event.ctrlKey || event.metaKey) || event.altKey || event.shiftKey) return null;
  const match = /^Digit([1-9])$/.exec(event.code);
  if (!match) return null;
  const digit = Number(match[1]);
  if (digit === 9) return orderedIds[orderedIds.length - 1] ?? null;
  return orderedIds[digit - 1] ?? null;
}

export function stepClassicTab(
  orderedIds: readonly string[],
  activeId: string | null,
  direction: 1 | -1,
): string | null {
  if (orderedIds.length === 0) return null;
  const index = activeId ? orderedIds.indexOf(activeId) : -1;
  if (index < 0) return orderedIds[direction === 1 ? 0 : orderedIds.length - 1] ?? null;
  return orderedIds[(index + direction + orderedIds.length) % orderedIds.length] ?? null;
}
