/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The presentation Ribbon's tab model: the base, contextual and Slide Master
 * tabs, their types and guards, the shape-kind message keys the contextual
 * tabs use, and normalising a stored tab to one the Ribbon shows.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import type {
  PresentationContextualRibbonTab,
} from "./presentation-ribbon-structure";
import type { ShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";

export const PRESENTATION_WEB_RIBBON_TABS = [
  "home",
  "insert",
  "draw",
  "design",
  "transitions",
  "animations",
  "slideshow",
  // B395: the owner's product scope is these permanent tabs (Help left the
  
  // (PPT-A001 / F007) is a contextual tab instead -- see
  // `presentationContextualRibbonTabs` -- shown while the recording studio is
  // open or a recording runs, and while a camera (Cameo) object is selected.
  "review",
  "view",
] as const;

/*
 * Animations > Advanced Animation follows desktop PowerPoint (PPT-A089):
 * Animation Pane, Trigger and Animation Painter sit on the Ribbon. The
 * Trigger menu is the only editor for `triggerBlockId`, which a deck
 * authored on the desktop carries through import, copy/paste and reuse.
 */

export type PresentationBaseRibbonTab = (typeof PRESENTATION_WEB_RIBBON_TABS)[number];

export const PRESENTATION_CONTEXTUAL_RIBBON_TABS = [
  "shape-format",
  "camera-format",
  "image-format",
  "graphic-format",
  "chart-design",
  "chart-format",
  "media-format",
  "media-playback",
  "smartart-design",
  "smartart-format",
  "table-design",
  "table-layout",
  "equation",
  "record",
] as const satisfies readonly PresentationContextualRibbonTab[];

/**
 * The Slide Master tab (PPT-B030): PowerPoint shows it first, and selects it,
 * only while Slide Master view is open.
 */
export const PRESENTATION_SLIDE_MASTER_RIBBON_TAB = "slide-master" as const;

const PRESENTATION_ALL_WEB_RIBBON_TABS = [
  ...PRESENTATION_WEB_RIBBON_TABS,
  ...PRESENTATION_CONTEXTUAL_RIBBON_TABS,
  PRESENTATION_SLIDE_MASTER_RIBBON_TAB,
] as const;

export type PresentationWebRibbonTab =
  | PresentationBaseRibbonTab
  | PresentationContextualRibbonTab
  | typeof PRESENTATION_SLIDE_MASTER_RIBBON_TAB;

export type PresentationRibbonTab = PresentationWebRibbonTab;

export function presentationShapeKindMessageKey(
  kind: ShapeKind,
): "presentation.shape.rectangle" | "presentation.shape.freeform" {
  return kind === "custom-geometry"
    ? "presentation.shape.freeform"
    : (`presentation.shape.${kind}` as "presentation.shape.rectangle");
}

export function isPresentationWebRibbonTab(
  tab: string,
): tab is PresentationWebRibbonTab {
  return (PRESENTATION_ALL_WEB_RIBBON_TABS as readonly string[]).includes(tab);
}

export function isPresentationBaseRibbonTab(
  tab: string,
): tab is PresentationBaseRibbonTab {
  return (PRESENTATION_WEB_RIBBON_TABS as readonly string[]).includes(tab);
}

export function isPresentationContextualRibbonTab(
  tab: string,
): tab is PresentationContextualRibbonTab {
  return (PRESENTATION_CONTEXTUAL_RIBBON_TABS as readonly string[]).includes(
    tab,
  );
}

export function normalizePresentationWebRibbonTab(
  tab: PresentationRibbonTab,
): PresentationWebRibbonTab {
  return isPresentationWebRibbonTab(tab) ? tab : "home";
}
