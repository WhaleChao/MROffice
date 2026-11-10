/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useMemo, useSyncExternalStore } from "react";
import type { OfficeWindowPageType } from "./windowCommands";

/**
 * Word's Synchronous Scrolling for two windows on one page. Each editor
 * window registers its scroll port (the element that scrolls its page) with
 * its page type and page id; while a session pairs two ports of the same
 * page, scrolling one moves the other by as much, down and across, as Word
 * and Excel do: each window keeps the place it was showing when the two
 * were paired (rows 1-20 beside rows 101-120 stay that far apart). A
 * proportional pairing pulled the other window to the same fraction of its
 * page on the first scroll, losing the part the user had set beside it, and
 * never moved it across (video audit S04).
 * Docs (components/console/editor/documentSynchronizedScrolling.ts) binds
 * these to "document".
 */

export interface OfficeSynchronizedScrollState {
  active: boolean;
  companionWindowId: string | null;
  supported: boolean;
}

export interface OfficeScrollPortRequest {
  element: HTMLElement;
  pageId: string;
  pageType: OfficeWindowPageType;
  windowId: string;
}

interface RegisteredScrollPort {
  element: HTMLElement;
  /** Where the port was at its last scroll event or when its session
   * started (0 before either). */
  lastTop: number;
  lastLeft: number;
  onScroll: () => void;
  pageKey: string;
  windowId: string;
}

interface ScrollSession {
  firstWindowId: string;
  secondWindowId: string;
}

const ports = new Map<string, RegisteredScrollPort>();
const sessions = new Map<string, ScrollSession>();
const listeners = new Set<() => void>();
let version = 0;

function pageKeyOf(pageType: OfficeWindowPageType, pageId: string): string {
  return `${pageType}:${pageId}`;
}

function emitChange(): void {
  version += 1;
  listeners.forEach((listener) => listener());
}

function sessionContains(session: ScrollSession, windowId: string): boolean {
  return (
    session.firstWindowId === windowId || session.secondWindowId === windowId
  );
}

function livePort(windowId: string, pageKey: string) {
  const port = ports.get(windowId);
  return port?.pageKey === pageKey && port.element.isConnected ? port : null;
}

function retireSessionForWindow(windowId: string): boolean {
  let retired = false;
  for (const [pageKey, session] of sessions) {
    if (!sessionContains(session, windowId)) continue;
    sessions.delete(pageKey);
    retired = true;
  }
  return retired;
}

function synchronizeFrom(sourceWindowId: string): void {
  const source = ports.get(sourceWindowId);
  if (!source) return;
  const { element } = source;
  const deltaTop = element.scrollTop - source.lastTop;
  const deltaLeft = element.scrollLeft - source.lastLeft;
  source.lastTop = element.scrollTop;
  source.lastLeft = element.scrollLeft;
  // The scroll this module wrote into a companion arrives here as no move.
  if (Math.abs(deltaTop) < 1 && Math.abs(deltaLeft) < 1) return;
  const session = sessions.get(source.pageKey);
  if (!session || !sessionContains(session, sourceWindowId)) return;
  const targetWindowId =
    session.firstWindowId === sourceWindowId
      ? session.secondWindowId
      : session.firstWindowId;
  const target = livePort(targetWindowId, source.pageKey);
  if (!target) {
    sessions.delete(source.pageKey);
    emitChange();
    return;
  }
  const into = target.element;
  if (Math.abs(deltaTop) >= 1) into.scrollTop += deltaTop;
  if (Math.abs(deltaLeft) >= 1) into.scrollLeft += deltaLeft;
  // What the companion took (its own ends clamp it) is where it now is.
  target.lastTop = into.scrollTop;
  target.lastLeft = into.scrollLeft;
}

/** Registers a window's scroll port; returns the unregister. */
export function registerOfficeScrollPort(
  request: OfficeScrollPortRequest,
): () => void {
  if (!request.pageId || !request.windowId) return () => undefined;
  const previous = ports.get(request.windowId);
  if (previous) {
    previous.element.removeEventListener("scroll", previous.onScroll);
    retireSessionForWindow(request.windowId);
  }
  // Not read here: a port registers in the layout effect of the editor that
  // mounts it, where reading its scroll position laid out the whole desk in
  
  // loaded). Nothing needs it before a session: registering retires this
  // window's sessions, a session starts by toggling (which reads both ports)
  // and a scroll outside a session only records where the port is.
  const port: RegisteredScrollPort = {
    element: request.element,
    lastTop: 0,
    lastLeft: 0,
    onScroll: () => synchronizeFrom(request.windowId),
    pageKey: pageKeyOf(request.pageType, request.pageId),
    windowId: request.windowId,
  };
  ports.set(request.windowId, port);
  request.element.addEventListener("scroll", port.onScroll, { passive: true });
  emitChange();
  return () => {
    const current = ports.get(request.windowId);
    if (current !== port) return;
    current.element.removeEventListener("scroll", current.onScroll);
    ports.delete(request.windowId);
    retireSessionForWindow(request.windowId);
    emitChange();
  };
}

export function officeSynchronizedScrollState(
  windowId: string,
  pageId: string,
  pageType: OfficeWindowPageType,
): OfficeSynchronizedScrollState {
  const pageKey = pageKeyOf(pageType, pageId);
  const current = livePort(windowId, pageKey);
  const companions = [...ports.values()].filter(
    (port) =>
      port.windowId !== windowId &&
      port.pageKey === pageKey &&
      port.element.isConnected,
  );
  const session = sessions.get(pageKey);
  const active = Boolean(
    current &&
    session &&
    sessionContains(session, windowId) &&
    livePort(
      session.firstWindowId === windowId
        ? session.secondWindowId
        : session.firstWindowId,
      pageKey,
    ),
  );
  return {
    active,
    companionWindowId: active
      ? session?.firstWindowId === windowId
        ? (session?.secondWindowId ?? null)
        : (session?.firstWindowId ?? null)
      : null,
    supported: Boolean(current && companions.length > 0),
  };
}

/** Turns the page's session on (with the first companion) or off. */
export function toggleOfficeSynchronizedScrolling(
  windowId: string,
  pageId: string,
  pageType: OfficeWindowPageType,
): boolean {
  const pageKey = pageKeyOf(pageType, pageId);
  const current = livePort(windowId, pageKey);
  if (!current) return false;
  const activeSession = sessions.get(pageKey);
  if (activeSession && sessionContains(activeSession, windowId)) {
    sessions.delete(pageKey);
    emitChange();
    return true;
  }
  const companion = [...ports.values()].find(
    (port) =>
      port.windowId !== windowId &&
      port.pageKey === pageKey &&
      port.element.isConnected,
  );
  if (!companion) return false;
  for (const port of [current, companion]) {
    port.lastTop = port.element.scrollTop;
    port.lastLeft = port.element.scrollLeft;
  }
  sessions.set(pageKey, {
    firstWindowId: windowId,
    secondWindowId: companion.windowId,
  });
  emitChange();
  return true;
}

/** For useSyncExternalStore: change notifications and a version snapshot. */
export function subscribeOfficeSynchronizedScrolling(
  listener: () => void,
): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

export function officeSynchronizedScrollingVersion(): number {
  return version;
}

function stateKey(state: OfficeSynchronizedScrollState): string {
  return `${state.active ? 1 : 0}|${state.supported ? 1 : 0}|${state.companionWindowId ?? ""}`;
}

const NO_SYNCHRONIZED_SCROLL_KEY = "0|0|";

/**
 * One window's synchronous-scrolling state, re-rendering that window only when
 * its own state changes. Every window used to subscribe to the module-wide
 * version, so each port registration re-rendered all of them; with about ten
 * windows open, mounting one more set off enough nested renders for React to
 * abandon it (Maximum update depth exceeded).
 */
export function useOfficeSynchronizedScrollState(
  windowId: string,
  pageId: string,
  pageType: OfficeWindowPageType,
): OfficeSynchronizedScrollState {
  const key = useSyncExternalStore(
    subscribeOfficeSynchronizedScrolling,
    () => stateKey(officeSynchronizedScrollState(windowId, pageId, pageType)),
    () => NO_SYNCHRONIZED_SCROLL_KEY,
  );
  return useMemo(() => {
    const [active, supported, companion] = key.split("|");
    return {
      active: active === "1",
      companionWindowId: companion || null,
      supported: supported === "1",
    };
  }, [key]);
}
