/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What the resume (/open/start) hands the notebook it opens, for this tab
 * only: which pages to open and where their windows stood on the homepage's
 * workbench, each page's zoom, and that the notebook's guided tour waits for
 * another visit (the visitor came to work on their files; the tour must not
 * cover them or take their focus). sessionStorage: a tab's own, gone with the
 * tab, never sent anywhere.
 */

import type { ClassicRect } from "./stash";

const SET_KEY = (notebookId: string) => `mrnote.classic.open-set:${notebookId}`;
const ZOOM_KEY = (pageId: string) => `mrnote.classic.zoom:${pageId}`;
const TOUR_KEY = "mrnote.onboarding.defer-notebook-tour";

export interface ClassicOpenEntry {
  readonly pageId: string;
  /** Relative to the workbench canvas the visitor used. */
  readonly frame?: ClassicRect;
  readonly zoom?: number;
}

export interface ClassicOpenSet {
  readonly canvas?: { readonly width: number; readonly height: number };
  readonly entries: readonly ClassicOpenEntry[];
}

function storage(): Storage | null {
  try {
    return typeof window === "undefined" ? null : window.sessionStorage;
  } catch {
    return null;
  }
}

function finite(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function rect(value: unknown): ClassicRect | undefined {
  const frame = value as Partial<ClassicRect> | null | undefined;
  return frame && finite(frame.x) && finite(frame.y) && finite(frame.width) && finite(frame.height)
    ? { x: frame.x, y: frame.y, width: frame.width, height: frame.height }
    : undefined;
}

export function handOverClassicOpenSet(notebookId: string, set: ClassicOpenSet): void {
  const store = storage();
  if (!store) return;
  try {
    store.setItem(SET_KEY(notebookId), JSON.stringify(set));
    for (const entry of set.entries) {
      if (finite(entry.zoom)) store.setItem(ZOOM_KEY(entry.pageId), String(entry.zoom));
    }
  } catch {
    // the windows then open on the cascade
  }
}

function take(key: string): string | null {
  const store = storage();
  if (!store) return null;
  try {
    const raw = store.getItem(key);
    store.removeItem(key);
    return raw;
  } catch {
    return null;
  }
}

/** Read once by the notebook, when it opens its windows. */
export function takeClassicOpenSet(notebookId: string): ClassicOpenSet | null {
  const raw = take(SET_KEY(notebookId));
  if (!raw) return null;
  try {
    const value = JSON.parse(raw) as { canvas?: unknown; entries?: unknown };
    const canvas = value.canvas as { width?: unknown; height?: unknown } | undefined;
    const entries = (Array.isArray(value.entries) ? value.entries : [])
      .filter((entry): entry is { pageId: string; frame?: unknown; zoom?: unknown } =>
        Boolean(entry) && typeof (entry as { pageId?: unknown }).pageId === "string")
      .map((entry) => ({
        pageId: entry.pageId,
        ...(rect(entry.frame) ? { frame: rect(entry.frame) } : {}),
        ...(finite(entry.zoom) ? { zoom: entry.zoom } : {}),
      }));
    return {
      ...(canvas && finite(canvas.width) && finite(canvas.height)
        ? { canvas: { width: canvas.width, height: canvas.height } }
        : {}),
      entries,
    };
  } catch {
    return null;
  }
}

/** The document zoom the page had on the homepage; read once by its editor. */
export function takeClassicOpenZoom(pageId: string): number | undefined {
  const raw = take(ZOOM_KEY(pageId));
  const zoom = raw === null ? Number.NaN : Number(raw);
  return Number.isFinite(zoom) && zoom > 0 ? zoom : undefined;
}

export function deferNotebookTour(notebookId: string): void {
  try {
    storage()?.setItem(TOUR_KEY, notebookId);
  } catch {
    // the tour then runs as usual
  }
}

export function notebookTourDeferred(notebookId: string): boolean {
  try {
    return storage()?.getItem(TOUR_KEY) === notebookId;
  } catch {
    return false;
  }
}
