/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  readOfficeWindowRect,
  readOfficeWindowViewport,
  type OfficeWindowRect,
  type OfficeWindowViewport,
} from "./windowViewport";
import type { OfficeWindowPageType } from "./windowCommands";

/**
 * Word's View Side by Side: the source window and a second window on the
 * same page, laid out as two halves of the visible canvas. Keyed by page
 * type; an editor marks its page root with the page id attribute
 * (`OFFICE_PAGE_ID_ATTRIBUTE` by default; Docs keeps data-document-page-id)
 * so the companion window can be found once it has mounted.
 */

export const OFFICE_PAGE_ID_ATTRIBUTE = "data-office-page-id";

export interface OfficeSideBySideLayout {
  primary: OfficeWindowRect;
  secondary: OfficeWindowRect;
}

const SIDE_BY_SIDE_GAP_PX = 16;
const SIDE_BY_SIDE_INSET_PX = 16;
const SIDE_BY_SIDE_MIN_WIDTH_PX = 380;
const SIDE_BY_SIDE_MIN_HEIGHT_PX = 360;

export function computeOfficeSideBySideLayout(
  viewport: OfficeWindowViewport,
): OfficeSideBySideLayout | null {
  if (
    !Number.isFinite(viewport.x) ||
    !Number.isFinite(viewport.y) ||
    !Number.isFinite(viewport.width) ||
    !Number.isFinite(viewport.height)
  ) {
    return null;
  }
  const usableWidth = viewport.width - SIDE_BY_SIDE_INSET_PX * 2;
  const usableHeight = viewport.height - SIDE_BY_SIDE_INSET_PX * 2;
  const width = Math.floor((usableWidth - SIDE_BY_SIDE_GAP_PX) / 2);
  const height = Math.floor(usableHeight);
  if (
    width < SIDE_BY_SIDE_MIN_WIDTH_PX ||
    height < SIDE_BY_SIDE_MIN_HEIGHT_PX
  ) {
    return null;
  }
  const x = viewport.x + SIDE_BY_SIDE_INSET_PX;
  const y = viewport.y + SIDE_BY_SIDE_INSET_PX;
  return {
    primary: { x, y, width, height },
    secondary: {
      x: x + width + SIDE_BY_SIDE_GAP_PX,
      y,
      width,
      height,
    },
  };
}

export function readOfficeSideBySideLayout(
  windowId: string,
): OfficeSideBySideLayout | null {
  const viewport = readOfficeWindowViewport(windowId);
  return viewport ? computeOfficeSideBySideLayout(viewport) : null;
}

/** Narrowest companion that still reads as a page beside its source. */
const COMPANION_MIN_WIDTH_PX = 480;

/**
 * Where a companion window opens next to its source (a compiled PDF beside
 * its Code window): in the free canvas to the source's right when the visible
 * part has room there, at up to `preferredWidth`; otherwise the two become
 * View Side by Side halves, which moves the source into the left half.
 * Null when the source is not on a measurable canvas or the canvas is too
 * small for halves: the caller falls back to its ordinary placement.
 */
export function placeOfficeCompanionWindow(
  actions: Pick<
    OfficeSideBySideActions,
    "moveWindow" | "resizeWindow" | "restoreWindow"
  >,
  sourceWindowId: string,
  preferredWidth: number,
): OfficeWindowRect | null {
  const viewport = readOfficeWindowViewport(sourceWindowId);
  const source = readOfficeWindowRect(sourceWindowId);
  if (!viewport || !source) return null;
  const x = source.x + source.width + SIDE_BY_SIDE_GAP_PX;
  const room = viewport.x + viewport.width - SIDE_BY_SIDE_INSET_PX - x;
  const y = Math.max(source.y, viewport.y + SIDE_BY_SIDE_INSET_PX);
  const height = Math.min(
    source.height,
    viewport.y + viewport.height - SIDE_BY_SIDE_INSET_PX - y,
  );
  if (room >= COMPANION_MIN_WIDTH_PX && height >= SIDE_BY_SIDE_MIN_HEIGHT_PX) {
    return { x, y, width: Math.min(preferredWidth, room), height };
  }
  const layout = computeOfficeSideBySideLayout(viewport);
  if (!layout) return null;
  actions.restoreWindow(sourceWindowId);
  actions.resizeWindow(
    sourceWindowId,
    layout.primary.width,
    layout.primary.height,
  );
  actions.moveWindow(sourceWindowId, layout.primary.x, layout.primary.y);
  return layout.secondary;
}

/** The windows whose editor shows this page, found by its id attribute. */
export function officeWindowIdsForPage(
  pageId: string,
  pageIdAttribute: string = OFFICE_PAGE_ID_ATTRIBUTE,
): string[] {
  if (typeof document === "undefined" || !pageId) return [];
  const ids = new Set<string>();
  for (const editor of document.querySelectorAll<HTMLElement>(
    `[${pageIdAttribute}]`,
  )) {
    if (editor.getAttribute(pageIdAttribute) !== pageId) continue;
    const shell = editor.closest<HTMLElement>("[data-window-id]");
    const id = shell?.dataset.windowId;
    if (id) ids.add(id);
  }
  return [...ids];
}

export interface OfficeSideBySideActions {
  focusWindow: (id: string) => void;
  moveWindow: (id: string, x: number, y: number) => void;
  openWindow: (payload: {
    type: "note";
    title: string;
    meta: { notebookId: string; pageId: string; pageType: string };
    force_new: true;
    beside_type: "note";
  }) => void;
  resizeWindow: (id: string, width: number, height: number) => void;
  restoreWindow: (id: string) => void;
}

export interface OfficeSideBySideRequest {
  actions: OfficeSideBySideActions;
  notebookId: string;
  pageId: string;
  pageType: OfficeWindowPageType;
  /** Defaults to OFFICE_PAGE_ID_ATTRIBUTE. */
  pageIdAttribute?: string;
  sourceWindowId: string;
  title: string;
  requestFrame?: (callback: FrameRequestCallback) => number;
  cancelFrame?: (handle: number) => void;
}

export interface OfficeSideBySideRun {
  accepted: boolean;
  cancel: () => void;
}

function arrangePair(
  actions: OfficeSideBySideActions,
  sourceWindowId: string,
  companionWindowId: string,
  layout: OfficeSideBySideLayout,
): void {
  actions.restoreWindow(sourceWindowId);
  actions.restoreWindow(companionWindowId);
  actions.resizeWindow(
    sourceWindowId,
    layout.primary.width,
    layout.primary.height,
  );
  actions.resizeWindow(
    companionWindowId,
    layout.secondary.width,
    layout.secondary.height,
  );
  actions.moveWindow(sourceWindowId, layout.primary.x, layout.primary.y);
  actions.moveWindow(companionWindowId, layout.secondary.x, layout.secondary.y);
  actions.focusWindow(sourceWindowId);
}

/**
 * Lays out the source window and a companion on the same page side by side,
 * opening the companion first when there is none (it is placed once its
 * editor has mounted, polled for up to 60 frames).
 */
export function startOfficeSideBySide(
  request: OfficeSideBySideRequest,
): OfficeSideBySideRun {
  const layout = readOfficeSideBySideLayout(request.sourceWindowId);
  if (!layout) return { accepted: false, cancel: () => undefined };
  const attribute = request.pageIdAttribute ?? OFFICE_PAGE_ID_ATTRIBUTE;
  const ids = officeWindowIdsForPage(request.pageId, attribute);
  const existingCompanion = ids.find(
    (candidate) => candidate !== request.sourceWindowId,
  );
  if (existingCompanion) {
    arrangePair(
      request.actions,
      request.sourceWindowId,
      existingCompanion,
      layout,
    );
    return { accepted: true, cancel: () => undefined };
  }

  const requestFrame = request.requestFrame ?? window.requestAnimationFrame;
  const cancelFrame = request.cancelFrame ?? window.cancelAnimationFrame;
  const beforeIds = new Set(ids);
  let frame = 0;
  let attemptsRemaining = 60;
  let cancelled = false;
  request.actions.openWindow({
    type: "note",
    title: request.title,
    meta: {
      notebookId: request.notebookId,
      pageId: request.pageId,
      pageType: request.pageType,
    },
    force_new: true,
    beside_type: "note",
  });
  const settle = () => {
    frame = 0;
    if (cancelled) return;
    const companion = officeWindowIdsForPage(request.pageId, attribute).find(
      (candidate) =>
        candidate !== request.sourceWindowId && !beforeIds.has(candidate),
    );
    if (companion) {
      arrangePair(request.actions, request.sourceWindowId, companion, layout);
      return;
    }
    attemptsRemaining -= 1;
    if (attemptsRemaining > 0) frame = requestFrame(settle);
  };
  frame = requestFrame(settle);
  return {
    accepted: true,
    cancel: () => {
      cancelled = true;
      if (frame) cancelFrame(frame);
      frame = 0;
    },
  };
}
