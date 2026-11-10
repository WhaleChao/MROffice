/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  WindowManagerActions,
  WindowSemanticState,
} from "../../components/notebook/WindowManager";

import {
  readOfficeWindowViewport,
  type OfficeWindowRect,
  type OfficeWindowViewport,
} from "./windowViewport";
import { CANVAS_COVERED_CHANGED_EVENT } from "../canvas-covered-edge";

/** The longest a layout waits for the sidebar to come back over the canvas. */
const SIDEBAR_RETURN_WAIT_MS = 400;









function afterWindowsRestored(
  actions: Pick<WindowManagerActions, "restoreWindow"> &
    Partial<Pick<WindowManagerActions, "readWindowFrame">>,
  ids: readonly string[],
  layOut: () => void,
): boolean {
  const maximized = ids.filter((id) => actions.readWindowFrame?.(id)?.maximized === true);
  if (!maximized.length || typeof window === "undefined") {
    layOut();
    return true;
  }
  for (const id of maximized) actions.restoreWindow(id);
  let done = false;
  const finish = () => {
    if (done) return;
    done = true;
    window.removeEventListener(CANVAS_COVERED_CHANGED_EVENT, settle);
    window.clearTimeout(timer);
    layOut();
  };
  // The sidebar publishes its reach after its own render: lay out on the
  // frame after it does.
  const settle = () => window.requestAnimationFrame(() => window.requestAnimationFrame(finish));
  window.addEventListener(CANVAS_COVERED_CHANGED_EVENT, settle);
  const timer = window.setTimeout(finish, SIDEBAR_RETURN_WAIT_MS);
  return false;
}

/**
 * Word's View > Window commands (New Window, Arrange All, Switch Windows),
 * keyed by page type so every editor's windows get the same commands over
 * windows of its own kind. Docs (components/console/editor/
 * documentWindowCommands.ts) binds these to "document".
 */

/** The page types an editor window can hold (WindowManager meta.pageType). */
export type OfficeWindowPageType =
  | "document"
  | "spreadsheet"
  | "presentation"
  | "pdf"
  | "code";

export interface OfficeWindowChoice {
  current: boolean;
  id: string;
  minimized: boolean;
  pageId: string;
  title: string;
}

const OFFICE_WINDOW_GAP_PX = 16;
const OFFICE_WINDOW_INSET_PX = 16;
const OFFICE_WINDOW_MIN_WIDTH_PX = 320;
const OFFICE_WINDOW_MIN_HEIGHT_PX = 260;
export const OFFICE_WINDOW_COMMAND_MAX_WINDOWS = 12;

/**
 * Whether a window is a page window of this type. A page window without a
 * page type is a document window (the legacy note window).
 */
export function isOfficeWindowOfPageType(
  windowState: WindowSemanticState,
  pageType: OfficeWindowPageType,
): windowState is WindowSemanticState & { meta: { pageId: string } } {
  const metaType = windowState.meta.pageType;
  return (
    (windowState.type === "note" || windowState.type === "guest_note") &&
    (pageType === "document"
      ? !metaType || metaType === "document"
      : metaType === pageType) &&
    typeof windowState.meta.pageId === "string" &&
    windowState.meta.pageId.length > 0
  );
}

/** Switch Windows: the type's windows, front-most first, at most 12. */
export function officeWindowChoices(
  windows: readonly WindowSemanticState[],
  currentWindowId: string,
  pageType: OfficeWindowPageType,
): OfficeWindowChoice[] {
  return windows
    .filter((windowState) => isOfficeWindowOfPageType(windowState, pageType))
    .sort(
      (left, right) =>
        right.zIndex - left.zIndex || left.id.localeCompare(right.id),
    )
    .slice(0, OFFICE_WINDOW_COMMAND_MAX_WINDOWS)
    .map((windowState) => ({
      current: windowState.id === currentWindowId,
      id: windowState.id,
      minimized: windowState.minimized,
      pageId: windowState.meta.pageId,
      title: windowState.title,
    }));
}

/** Arrange All's tiles: a near-square grid, or null when they would not fit. */
export function computeOfficeArrangeAllLayout(
  viewport: OfficeWindowViewport,
  count: number,
): OfficeWindowRect[] | null {
  if (
    !Number.isInteger(count) ||
    count < 1 ||
    count > OFFICE_WINDOW_COMMAND_MAX_WINDOWS ||
    !Number.isFinite(viewport.x) ||
    !Number.isFinite(viewport.y) ||
    !Number.isFinite(viewport.width) ||
    !Number.isFinite(viewport.height)
  ) {
    return null;
  }
  const columns = Math.ceil(Math.sqrt(count));
  const rows = Math.ceil(count / columns);
  const usableWidth =
    viewport.width -
    OFFICE_WINDOW_INSET_PX * 2 -
    OFFICE_WINDOW_GAP_PX * (columns - 1);
  const usableHeight =
    viewport.height -
    OFFICE_WINDOW_INSET_PX * 2 -
    OFFICE_WINDOW_GAP_PX * (rows - 1);
  const width = Math.floor(usableWidth / columns);
  const height = Math.floor(usableHeight / rows);
  if (
    width < OFFICE_WINDOW_MIN_WIDTH_PX ||
    height < OFFICE_WINDOW_MIN_HEIGHT_PX
  ) {
    return null;
  }
  const x = viewport.x + OFFICE_WINDOW_INSET_PX;
  const y = viewport.y + OFFICE_WINDOW_INSET_PX;
  return Array.from({ length: count }, (_, index) => ({
    x: x + (index % columns) * (width + OFFICE_WINDOW_GAP_PX),
    y: y + Math.floor(index / columns) * (height + OFFICE_WINDOW_GAP_PX),
    width,
    height,
  }));
}







export function openNewOfficeWindow(
  actions: Pick<WindowManagerActions, "openWindow"> &
    Partial<Pick<WindowManagerActions, "readWindowFrame" | "restoreWindow">>,
  request: {
    notebookId: string;
    pageId: string;
    pageType: OfficeWindowPageType;
    title: string;
    /** The window the command was given in. */
    sourceWindowId?: string;
  },
): boolean {
  if (!request.notebookId || !request.pageId) return false;
  if (
    request.sourceWindowId &&
    actions.readWindowFrame?.(request.sourceWindowId)?.maximized === true
  ) {
    actions.restoreWindow?.(request.sourceWindowId);
  }
  actions.openWindow({
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
  return true;
}

/** Arrange All: tiles the type's windows over the source window's canvas. */
export function arrangeAllOfficeWindows(
  actions: Pick<
    WindowManagerActions,
    | "focusWindow"
    | "moveWindow"
    | "readWindows"
    | "resizeWindow"
    | "restoreWindow"
  > &
    Partial<Pick<WindowManagerActions, "readWindowFrame">>,
  sourceWindowId: string,
  pageType: OfficeWindowPageType,
): boolean {
  const choices = officeWindowChoices(
    actions.readWindows(),
    sourceWindowId,
    pageType,
  );
  if (!choices.length) return false;
  const focusId = choices.some((choice) => choice.id === sourceWindowId)
    ? sourceWindowId
    : choices[0]?.id;
  let arranged = false;
  const now = afterWindowsRestored(actions, choices.map((choice) => choice.id), () => {
    const viewport = readOfficeWindowViewport(sourceWindowId);
    const layout = viewport
      ? computeOfficeArrangeAllLayout(viewport, choices.length)
      : null;
    if (!layout) return;
    choices.forEach((choice, index) => {
      const rect = layout[index];
      if (!rect) return;
      actions.restoreWindow(choice.id);
      actions.resizeWindow(choice.id, rect.width, rect.height);
      actions.moveWindow(choice.id, rect.x, rect.y);
    });
    if (focusId) actions.focusWindow(focusId);
    arranged = Boolean(focusId);
  });
  // Laid out now, or (from a maximized window) a moment later.
  return now ? arranged : true;
}

/** Switch Windows: brings one of the type's windows to the front. */
export function switchToOfficeWindow(
  actions: Pick<
    WindowManagerActions,
    "focusWindow" | "readWindows" | "restoreWindow"
  >,
  targetWindowId: string,
  pageType: OfficeWindowPageType,
): boolean {
  const target = actions
    .readWindows()
    .find(
      (windowState) =>
        windowState.id === targetWindowId &&
        isOfficeWindowOfPageType(windowState, pageType),
    );
  if (!target) return false;
  actions.restoreWindow(target.id);
  actions.focusWindow(target.id);
  return true;
}

/** Cascade's diagonal step and window size share, as Office staggers its windows. */
const OFFICE_CASCADE_STEP_PX = 32;

/** Cascade: the type's windows stacked diagonally, front-most last (on top). */
export function computeOfficeCascadeLayout(
  viewport: OfficeWindowViewport,
  count: number,
): OfficeWindowRect[] | null {
  if (
    !Number.isInteger(count) ||
    count < 1 ||
    count > OFFICE_WINDOW_COMMAND_MAX_WINDOWS ||
    ![viewport.x, viewport.y, viewport.width, viewport.height].every(Number.isFinite)
  ) {
    return null;
  }
  const usableWidth = viewport.width - OFFICE_WINDOW_INSET_PX * 2;
  const usableHeight = viewport.height - OFFICE_WINDOW_INSET_PX * 2;
  const width = Math.floor(usableWidth - OFFICE_CASCADE_STEP_PX * (count - 1));
  const height = Math.floor(usableHeight - OFFICE_CASCADE_STEP_PX * (count - 1));
  if (width < OFFICE_WINDOW_MIN_WIDTH_PX || height < OFFICE_WINDOW_MIN_HEIGHT_PX) {
    return null;
  }
  return Array.from({ length: count }, (_, index) => ({
    x: viewport.x + OFFICE_WINDOW_INSET_PX + index * OFFICE_CASCADE_STEP_PX,
    y: viewport.y + OFFICE_WINDOW_INSET_PX + index * OFFICE_CASCADE_STEP_PX,
    width,
    height,
  }));
}

/** View ▸ Window ▸ Cascade over the type's windows; the source ends on top. */
export function cascadeOfficeWindows(
  actions: Pick<
    WindowManagerActions,
    | "focusWindow"
    | "moveWindow"
    | "readWindows"
    | "resizeWindow"
    | "restoreWindow"
  > &
    Partial<Pick<WindowManagerActions, "readWindowFrame">>,
  sourceWindowId: string,
  pageType: OfficeWindowPageType,
): boolean {
  // Back to front, so the window the user is in lands last, on top.
  const all = officeWindowChoices(actions.readWindows(), sourceWindowId, pageType);
  const choices = [
    ...all.filter((choice) => !choice.current).reverse(),
    ...all.filter((choice) => choice.current),
  ];
  if (!choices.length) return false;
  let cascaded = false;
  const now = afterWindowsRestored(actions, choices.map((choice) => choice.id), () => {
    const viewport = readOfficeWindowViewport(sourceWindowId);
    const layout = viewport ? computeOfficeCascadeLayout(viewport, choices.length) : null;
    if (!layout) return;
    choices.forEach((choice, index) => {
      const rect = layout[index];
      if (!rect) return;
      actions.restoreWindow(choice.id);
      actions.resizeWindow(choice.id, rect.width, rect.height);
      actions.moveWindow(choice.id, rect.x, rect.y);
      actions.focusWindow(choice.id);
    });
    cascaded = true;
  });
  return now ? cascaded : true;
}
