"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  bringOfficeTaskPaneForward,
  reconcileOfficeTaskPaneDock,
  useOfficeTaskPaneDock,
  type OfficeTaskPaneDock,
  type OfficeTaskPaneDockState,
} from "@mrnote/core/web/components/office-ribbon/officeTaskPaneDock.ts";

/**
 * The slide editor's right-hand task pane dock.
 *
 * PowerPoint docks Format Shape / Format Background, the Selection Pane, the
 * Animation Pane, the Accessibility pane, the Alt Text pane, Find (the
 * results list Word's Navigation pane keeps) and Designer's Design Ideas in
 * ONE place: the right edge of
 * the editing area (Microsoft Support, "Manage pane layout in Office apps").
 * Opening a second pane does not lay it over the first. The dock shows one
 * pane at a time, and every open pane keeps a tab there until its own X
 * closes it; the pane just opened is the one in front (Microsoft Q&A,
 * "PowerPoint 2019 - View both Animation and Selection Pane at the same
 * time": both are "docked" to the right and only one is displayed at a
 * time). The editor used to float three of these as cards of their own at
 * three different spots, so two open at once covered each other.
 *
 * This is the slide editor's adapter onto the shared dock model
 * (components/office-ribbon/officeTaskPaneDock.ts, which the document editor
 * uses too): its pane ids, their title ids and its tab test ids. The model
 * remembers the order panes were opened in (the tab order) and which was
 * brought forward last (the front), and hands each pane its tab strip. It
 * owns no pane's open state; a pane that closes simply drops out, and the
 * one brought forward before it comes back.
 */

export const PRESENTATION_TASK_PANE_IDS = [
  "format",
  "background",
  "selection",
  "animation",
  "accessibility",
  "altText",
  "find",
  "designer",
  "reuse",
  /** Version history (the shared VersionHistoryPanel, docked like the rest). */
  "versionHistory",
  /** Review > Thesaurus (the shared OfficeThesaurusPane). */
  "thesaurus",
] as const;

export type PresentationTaskPaneId =
  (typeof PRESENTATION_TASK_PANE_IDS)[number];

export type PresentationTaskPaneFlags = Readonly<
  Record<PresentationTaskPaneId, boolean>
>;

/** Each pane's title element id: its own (selected) tab in the strip, and
 * what the pane is labelled by. These are the ids the panes carried before
 * they joined the dock. */
export const PRESENTATION_TASK_PANE_TITLE_IDS: Readonly<
  Record<PresentationTaskPaneId, string>
> = {
  format: "presentation-block-inspector-title",
  background: "presentation-background-pane-title",
  selection: "presentation-selection-pane-heading",
  animation: "presentation-animation-pane-tab",
  accessibility: "presentation-accessibility-panel-title",
  // The shared Alt Text pane's own title id (`<prefix>-pane-title`).
  altText: "presentation-alt-text-pane-title",
  find: "presentation-find-pane-title",
  designer: "presentation-design-ideas-pane-title",
  reuse: "presentation-reuse-slides-pane-title",
  versionHistory: "document-version-history-pane-title",
  thesaurus: "presentation-thesaurus-pane-title",
};

export type PresentationTaskPaneDockState =
  OfficeTaskPaneDockState<PresentationTaskPaneId>;

/**
 * Brings the remembered order up to date with which panes are open now. A
 * pane that has just opened joins at the end of both lists (it is the one in
 * front); a pane that has closed leaves both. Returns `state` itself when
 * nothing changed, so a render can call it on every pass.
 */
export function reconcilePresentationTaskPaneDock(
  state: PresentationTaskPaneDockState,
  open: PresentationTaskPaneFlags,
): PresentationTaskPaneDockState {
  return reconcileOfficeTaskPaneDock(state, PRESENTATION_TASK_PANE_IDS, open);
}

export function bringPresentationTaskPaneForward(
  state: PresentationTaskPaneDockState,
  id: PresentationTaskPaneId,
): PresentationTaskPaneDockState {
  return bringOfficeTaskPaneForward(state, id);
}

export type PresentationTaskPaneDock = OfficeTaskPaneDock<PresentationTaskPaneId>;

const presentationTaskPaneTabTestId = (id: PresentationTaskPaneId) =>
  `presentation-task-pane-tab-${id}`;

/**
 * `open` is each pane's own open state. `showing` narrows it to the panes
 * that have something to show right now: Format Shape stays open while no
 * object is selected (PowerPoint keeps its tab), but it has no body then, so
 * the pane before it stays in front instead of an empty dock.
 */
export function usePresentationTaskPaneDock(
  open: PresentationTaskPaneFlags,
  showing: PresentationTaskPaneFlags,
  labels: Readonly<Record<PresentationTaskPaneId, string>>,
  /** A tab of the strip brought its pane forward: the pane that held the
   * focus is now hidden, so the host moves the focus to the new front pane's
   * own tab. */
  onTabSelected?: (id: PresentationTaskPaneId) => void,
): PresentationTaskPaneDock {
  return useOfficeTaskPaneDock({
    order: PRESENTATION_TASK_PANE_IDS,
    open,
    showing,
    labels,
    tabTestId: presentationTaskPaneTabTestId,
    onTabSelected,
  });
}
