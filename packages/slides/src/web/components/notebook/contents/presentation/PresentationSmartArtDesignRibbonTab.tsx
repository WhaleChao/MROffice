"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The SmartArt Design contextual tab of the presentation Ribbon, drawn from
 * the window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import {
  OfficeSmartArtDesignTab,
  type OfficeSmartArtDesignTabDiagram,
  type OfficeSmartArtDesignTabLabels,
  type OfficeSmartArtDesignTabPreserved,
} from "@mrnote/core/web/components/office-ribbon/OfficeSmartArtDesignTab.tsx";
import { PRESENTATION_SMARTART_DESIGN_TAB_CLASSES } from "../../../../lib/presentation/window-constants";
import { PresentationShapeGraphic } from "./PresentationShapeGraphic";
import type { PresentationShapeBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { OFFICE_SMARTART_POWERPOINT_COLOR_GROUPS, type SmartArtColorTheme } from "@mrnote/core/web/lib/office-smartart/colors.ts";
import type { OfficeSmartArtPlaceholders } from "@mrnote/core/web/lib/office-smartart/catalog.ts";
import type { RefObject } from "react";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationSmartArtDesignRibbonTabCtl {
  smartArtPlaceholders: OfficeSmartArtPlaceholders;
  smartArtDesignLabels: OfficeSmartArtDesignTabLabels;
  smartArtTextPaneTriggerRef: RefObject<HTMLButtonElement | null>;
  activeSlideBounds: { width: number; height: number; };
  themeAccentColors: string[];
  deckSmartArtColorTheme: SmartArtColorTheme | null;
  presentationSmartArtPreserved: OfficeSmartArtDesignTabPreserved | null;
  presentationSmartArtDiagram: OfficeSmartArtDesignTabDiagram | null;
}

export function PresentationSmartArtDesignRibbonTab({ ctl }: { ctl: PresentationSmartArtDesignRibbonTabCtl }) {
  const {
    smartArtPlaceholders,
    smartArtDesignLabels,
    smartArtTextPaneTriggerRef,
    activeSlideBounds,
    themeAccentColors,
    deckSmartArtColorTheme,
    presentationSmartArtPreserved,
    presentationSmartArtDiagram,
  } = ctl;
  return (
                  <OfficeSmartArtDesignTab
                    testIdPrefix="presentation"
                    application="presentation"
                    classes={PRESENTATION_SMARTART_DESIGN_TAB_CLASSES}
                    controlIds={{
                      colors: "PPTCTL.CONTEXT.SMARTART.DESIGN.COLORS",
                    }}
                    textPaneTriggerRef={smartArtTextPaneTriggerRef}
                    placeholders={smartArtPlaceholders}
                    previewBounds={activeSlideBounds}
                    renderStyleShape={(node) => (
                      <PresentationShapeGraphic
                        block={node as unknown as PresentationShapeBlock}
                      />
                    )}
                    labels={smartArtDesignLabels}
                    preserved={presentationSmartArtPreserved}
                    diagram={presentationSmartArtDiagram}
                    // Change Colors resolves Office's variations against the
                    // deck's own theme, as PowerPoint does, and offers
                    // PowerPoint's whole gallery (PPT-D008).
                    colorTheme={deckSmartArtColorTheme ?? { accents: themeAccentColors }}
                    colorGroups={OFFICE_SMARTART_POWERPOINT_COLOR_GROUPS}
                  />
  );
}
