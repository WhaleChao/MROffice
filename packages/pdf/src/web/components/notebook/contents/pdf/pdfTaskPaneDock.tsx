"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useMemo, type ReactNode } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  OfficeDockedPane,
  OfficeTaskPaneDockHost,
} from "@mrnote/core/web/components/office-ribbon/OfficeTaskPaneDockHost.tsx";
import {
  keepOfficeTaskPaneRoom,
  useOfficeTaskPaneCover,
  type OfficeTaskPaneCover,
} from "@mrnote/core/web/components/office-ribbon/officeTaskPaneOverlay.ts";
import {
  useOfficeTaskPaneDockRegistry,
  type OfficeTaskPaneDockRegistry,
} from "@mrnote/core/web/components/office-ribbon/officeTaskPaneDock.ts";

import type { PdfNavigationPanel } from "./PdfNavigationRail";
import styles from "./PdfTaskPaneDock.module.css";

/**
 * The PDF reader's task pane dock: its adapter onto the shared dock model
 * (components/office-ribbon/officeTaskPaneDock.ts) and host
 * (OfficeTaskPaneDockHost), as Docs' is documentTaskPaneDock.tsx.
 *
 * Page Thumbnails, Bookmarks (the outline), Search results and Comments are
 * the reader's navigation panes; Acrobat and Word's Navigation pane dock them
 * at the LEFT edge of the window, beside the navigation rail that opens them.
 * The rail stays the switcher (one pane at a time, PdfNavigationRail); the
 * dock gives each pane the shared task pane chrome and its place.
 */

export type PdfTaskPaneEdge = "left" | "right";

/** The navigation panes, and the format panes of a selected object
 * (Format Shape, Format Chart): Office and ONLYOFFICE set those at the
 * right edge, beside what they format. */
export type PdfTaskPaneId = PdfNavigationPanel | "format-shape" | "format-chart";

export const PDF_TASK_PANE_EDGE: Readonly<Record<PdfTaskPaneId, PdfTaskPaneEdge>> = {
  thumbnails: "left",
  outline: "left",
  search: "left",
  annotations: "left",
  "format-shape": "right",
  "format-chart": "right",
};

export interface PdfTaskPaneDocks {
  readonly left: OfficeTaskPaneDockRegistry;
  readonly right: OfficeTaskPaneDockRegistry;
}

const pdfTaskPaneTabTestId = (id: string) => `pdf-task-pane-tab-${id}`;

export function usePdfTaskPaneDocks(): PdfTaskPaneDocks {
  const t = useTranslations("console-notebooks");
  // The strip's accessible name is the one the slide and document docks read.
  const tabsLabel = t("presentation.taskPanes.tabs");
  const options = useMemo(
    () => ({ tabTestId: pdfTaskPaneTabTestId, tabsLabel }),
    [tabsLabel],
  );
  const left = useOfficeTaskPaneDockRegistry(options);
  const right = useOfficeTaskPaneDockRegistry(options);
  return useMemo(() => ({ left, right }), [left, right]);
}

/** The custom property and attribute through which the pages keep room
 * for the cards (PdfTaskPaneDock.module.css). */
const PDF_TASK_PANE_ROOM = {
  variable: "--pdf-task-pane-room",
  attribute: "data-task-pane-room",
} as const;

/** The anchor the reader's panes hang from, between the rail and the pages:
 * each pane floats over the pages' left edge (PdfTaskPaneDock.module.css).
 * It publishes what the cards over the pages cover -- this dock's, and the
 * page window's Comments card at the right -- and keeps the pages' room for
 * them (office-ribbon/officeTaskPaneOverlay.ts). */
export function PdfTaskPaneDockHost({
  docks,
}: {
  readonly docks: PdfTaskPaneDocks;
}): React.JSX.Element {
  const body = docks.left.host?.parentElement ?? null;
  const root = body?.closest<HTMLElement>(".pdf-page-window__body") ?? body;
  const pages = body?.querySelector<HTMLElement>(":scope > .pdf-viewer__pages") ?? null;
  const onChange = useCallback(
    (cover: OfficeTaskPaneCover) => {
      if (!pages) return;
      // The Comments card floats over the pages' right edge as Docs' does:
      
      
      // reader's own panes at the left keep the pages' room.
      keepOfficeTaskPaneRoom({
        scroller: pages,
        viewport: pages,
        cover: { left: cover.left, right: 0 },
        holder: pages,
        ...PDF_TASK_PANE_ROOM,
      });
    },
    [pages],
  );
  useOfficeTaskPaneCover({ root, viewport: pages, onChange });
  return (
    <>
      <OfficeTaskPaneDockHost
        docks={docks}
        edge="left"
        className={styles.left}
        testId="pdf-task-pane-dock-left"
        overlay
      />
      {/* The format panes float over the pages' right edge, as the
          Comments card does: the pages keep their place. */}
      <OfficeTaskPaneDockHost
        docks={docks}
        edge="right"
        className={styles.right}
        testId="pdf-task-pane-dock-right"
        overlay
      />
    </>
  );
}

/** One reader pane in its dock. Without docks (a pane's own unit test) it
 * renders where it is. */
export function PdfDockedPane({
  docks,
  pane,
  children,
}: {
  readonly docks: PdfTaskPaneDocks | null;
  readonly pane: PdfTaskPaneId;
  readonly children: ReactNode;
}): React.JSX.Element | null {
  return (
    <OfficeDockedPane docks={docks} edges={PDF_TASK_PANE_EDGE} pane={pane}>
      {children}
    </OfficeDockedPane>
  );
}

/**
 * The target-size and focus utilities every PDF pane control carries (the
 * reader's 44px hit-area contract); the shared chrome draws the visible X.
 */
export const PDF_TASK_PANE_CLOSE_CLASS =
  "min-h-11 min-w-11 focus-visible:outline-none focus-visible:ring-2 " +
  "focus-visible:ring-ring focus-visible:ring-offset-1";

/**
 * The classes every docked reader pane adds to its OfficeTaskPane root: its
 * reader class (width, dark tokens). Its entrance and exit are the shared
 * pane's own (office-ribbon/officeTaskPaneMotion.ts).
 */
export function pdfTaskPaneClassName(readerClass: string): string {
  return `${readerClass} pdf-viewer__task-pane`;
}

/** A reader pane's scrolling content under the chrome's fixed title row. */
export const PDF_TASK_PANE_BODY_CLASS =
  `${officeDialogParts.paneBody} pdf-viewer__task-pane-body`;
