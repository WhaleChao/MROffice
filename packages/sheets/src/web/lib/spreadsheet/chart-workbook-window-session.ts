/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ReactNode } from "react";

export interface ChartWorkbookWindowSession {
  content: (windowId: string) => ReactNode;
  beforeClose: () => boolean;
  close: () => void;
  /** The canvas window whose editor holds this session (a document or a
   * deck). While the session lives, that window stays mounted when it is
   * minimized: unmounted, its editor ended the session and closed the data
   * window with it. */
  hostWindowId?: string;
}
const sessions = new Map<
  string,
  { owner: symbol; value: ChartWorkbookWindowSession }
>();
const listeners = new Map<string, Set<() => void>>();
const activators = new Map<string, () => void>();
const hostListeners = new Set<() => void>();
let hostWindowIds: ReadonlySet<string> = new Set();
const notify = (id: string) => {
  listeners.get(id)?.forEach((listener) => listener());
  const next = new Set<string>();
  for (const { value } of sessions.values()) {
    if (value.hostWindowId) next.add(value.hostWindowId);
  }
  if (next.size === hostWindowIds.size && [...next].every((hostId) => hostWindowIds.has(hostId))) return;
  hostWindowIds = next;
  hostListeners.forEach((listener) => listener());
};

/** The windows holding a live data window session (see `hostWindowId`). */
export function chartWorkbookHostWindowIds(): ReadonlySet<string> {
  return hostWindowIds;
}
export function subscribeChartWorkbookHostWindows(listener: () => void) {
  hostListeners.add(listener);
  return () => {
    hostListeners.delete(listener);
  };
}

export function publishChartWorkbookWindow(
  id: string,
  owner: symbol,
  value: ChartWorkbookWindowSession,
) {
  if (sessions.has(id) && sessions.get(id)?.owner !== owner)
    throw Error("chart_workbook_session_owner_mismatch");
  sessions.set(id, { owner, value });
  notify(id);
}
export function releaseChartWorkbookWindow(id: string, owner: symbol) {
  if (sessions.get(id)?.owner !== owner) return;
  sessions.delete(id);
  activators.delete(id);
  notify(id);
}
export function readChartWorkbookWindow(id: string) {
  return sessions.get(id)?.value ?? null;
}
export function subscribeChartWorkbookWindow(id: string, listener: () => void) {
  const set = listeners.get(id) ?? new Set();
  set.add(listener);
  listeners.set(id, set);
  return () => {
    set.delete(listener);
    if (!set.size) listeners.delete(id);
  };
}
export function registerChartWorkbookWindowActivator(
  id: string,
  activate: () => void,
) {
  activators.set(id, activate);
  return () => {
    if (activators.get(id) === activate) activators.delete(id);
  };
}
export function activateChartWorkbookWindow(id: string): boolean {
  const activate = activators.get(id);
  if (!sessions.has(id) || !activate) return false;
  activate();
  return true;
}
