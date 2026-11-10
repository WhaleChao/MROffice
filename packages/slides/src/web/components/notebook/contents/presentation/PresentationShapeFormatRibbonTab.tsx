"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Shape Format contextual tab of the presentation Ribbon: the shared
 * Office Shape Format tab driven by the presentation adapter.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { OfficeShapeFormatTab, type OfficeShapeFormatAdapter } from "@mrnote/core/web/components/office-ribbon/OfficeShapeFormatTab.tsx";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationShapeFormatRibbonTabCtl {
  presentationShapeFormatAdapter: OfficeShapeFormatAdapter;
}

export function PresentationShapeFormatRibbonTab({ ctl }: { ctl: PresentationShapeFormatRibbonTabCtl }) {
  const {
    presentationShapeFormatAdapter,
  } = ctl;
  return (
                <OfficeShapeFormatTab
                  adapter={presentationShapeFormatAdapter}
                  testIdPrefix="presentation-ribbon-shape"
                  menuClassName="docs-cupertino-menu"
                  ids={{
                    changeShape: "presentation-ribbon-shape-kind",
                    textBox: "presentation-ribbon-shape-format-add-textbox",
                    styles: "presentation-ribbon-shape-quick-styles",
                    fillSplit: "presentation-ribbon-shape-fill-cell",
                    fillPrimary: "presentation-ribbon-shape-fill",
                    fillTrigger: "presentation-ribbon-shape-fill-menu",
                    fill: "presentation-shape-format-fill",
                    outlineSplit: "presentation-ribbon-shape-stroke-cell",
                    outlinePrimary: "presentation-ribbon-shape-stroke",
                    outlineTrigger: "presentation-ribbon-shape-stroke-menu",
                    outline: "presentation-shape-format-stroke",
                    outlineDashes: "presentation-ribbon-shape-outline-style",
                    arrows: "presentation-ribbon-shape-outline-arrows",
                    lineEnds: "presentation-shape",
                    effect: "presentation-ribbon-shape-effect",
                    altText: "presentation-ribbon-shape-alt-text",
                    arrangeGroup: "presentation-ribbon-shape-arrange-group",
                    layerSplit: "presentation-ribbon-shape-bring-split",
                    layer: "presentation-ribbon-shape-bring-primary",
                    layerMenu: "presentation-ribbon-shape-bring-menu",
                    sendSplit: "presentation-ribbon-shape-send-split",
                    send: "presentation-ribbon-shape-send-primary",
                    sendMenu: "presentation-ribbon-shape-send-menu",
                    bringForwardItem: "presentation-ribbon-bring-forward",
                    bringToFrontItem: "presentation-ribbon-bring-to-front",
                    sendBackwardItem: "presentation-ribbon-send-backward",
                    sendToBackItem: "presentation-ribbon-send-to-back",
                    group: "presentation-ribbon-shape-group-menu",
                    groupItem: "presentation-group",
                    ungroupItem: "presentation-ungroup",
                    align: "presentation-ribbon-shape-align-menu",
                    alignToTarget: "presentation-align-to-slide",
                    alignLeft: "presentation-align-left",
                    alignCenter: "presentation-align-hcenter",
                    alignRight: "presentation-align-right",
                    alignTop: "presentation-align-top",
                    alignMiddle: "presentation-align-vcenter",
                    alignBottom: "presentation-align-bottom",
                    alignDistributeHorizontal: "presentation-distribute-horizontal",
                    alignDistributeVertical: "presentation-distribute-vertical",
                    alignMenu: "presentation-align-toolbar",
                    selectionCount: "presentation-align-selection-count",
                    rotate: "presentation-ribbon-shape-rotate-menu",
                    rotateRight90: "presentation-rotate-right",
                    rotateLeft90: "presentation-rotate-left",
                    flipHorizontal: "presentation-flip-horizontal",
                    flipVertical: "presentation-flip-vertical",
                    sizeGroup: "presentation-ribbon-shape-size-group",
                    sizeRow: "presentation-ribbon-shape-size-row",
                    size: "presentation-ribbon-shape",
                    lockAspect: "presentation-ribbon-shape-lock-aspect",
                  }}
                />
  );
}
