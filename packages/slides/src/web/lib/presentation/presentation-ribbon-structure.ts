/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PresentationContextualRibbonTab =
  | "shape-format"
  | "camera-format"
  | "image-format"
  | "graphic-format"
  | "chart-design"
  | "chart-format"
  | "media-format"
  | "media-playback"
  | "smartart-design"
  | "smartart-format"
  | "table-design"
  | "table-layout"
  /** The Equation tab (公式), over the text box's own tabs (PPT-A052). */
  | "equation"
  /**
   * The Record tab (录制). Not one of the ten permanent tabs (B395): it shows
   * while the recording studio is open or a recording runs, and while a
   * camera (Cameo) object is selected.
   */
  | "record";

export type PresentationContextualSelectionType =
  | "table"
  | "chart"
  | "media"
  | "smartart"
  | "graphic"
  | "image"
  | "camera"
  | "shape";

export type PresentationContextualMediaKind = "audio" | "video" | "iframe";

/**
 * The contextual tabs a selection opens, in ribbon order.
 *
 * Media: desktop PowerPoint opens "Video Format" + "Playback" for a video and
 * "Audio Format" + "Playback" for audio, the Format tab first (PPT-A006,
 * PPT-A007). PowerPoint for the web folds a video into one "Video" tab.
 *
 * SmartArt: desktop PowerPoint opens "SmartArt Design" + "Format", the
 * design tab first (PPT-A010); Format formats the shapes selected inside it.
 */
export function presentationContextualRibbonTabs(
  selection: PresentationContextualSelectionType | null,
  equation = false,
  record = false,
): readonly PresentationContextualRibbonTab[] {
  // An equation selected (or open in the editor) inside a text box adds
  // PowerPoint's Equation tab after the box's own Shape Format.
  const objectTabs = presentationObjectRibbonTabs(selection);
  const tabs: PresentationContextualRibbonTab[] = equation
    ? [...objectTabs, "equation"]
    : [...objectTabs];
  // Record is a mode tab, last: the recording studio (or a running
  // recording) shows it, and so does a selected camera object, whose
  // narration and Cameo commands live there.
  if (record || selection === "camera") tabs.push("record");
  return tabs;
}

function presentationObjectRibbonTabs(
  selection: PresentationContextualSelectionType | null,
): readonly PresentationContextualRibbonTab[] {
  switch (selection) {
    case "table":
      return ["table-design", "table-layout"];
    case "chart":
      
      return ["chart-design", "chart-format"];
    case "media":
      return ["media-format", "media-playback"];
    case "smartart":
      return ["smartart-design", "smartart-format"];
    case "camera":
      return ["camera-format"];
    case "graphic":
      return ["graphic-format"];
    case "image":
      return ["image-format"];
    case "shape":
      return ["shape-format"];
    default:
      return [];
  }
}

/**
 * The tab a NEW selection lands on. Audio lands on Playback (PowerPoint
 * switches to Playback when audio is inserted); every other family on its
 * first tab.
 */
export function presentationContextualRibbonDefaultTab(
  selection: PresentationContextualSelectionType | null,
  mediaKind: PresentationContextualMediaKind | null,
  equation = false,
): PresentationContextualRibbonTab | null {
  // PowerPoint switches to the Equation tab when an equation is inserted or
  // clicked into.
  if (equation) return "equation";
  if (selection === "media" && mediaKind === "audio") return "media-playback";
  return presentationContextualRibbonTabs(selection)[0] ?? null;
}

/**
 * The message key that names each contextual tab, verbatim. Desktop names
 * every object tab after what it formats -- "Shape Format", "Picture Format",
 * "Graphics Format", "Video Format", "Audio Format" (PPT-A003..A007); the
 * names are used as they stand, never with the "Format" word stripped.
 */
export function presentationContextualRibbonTabLabelKey(
  tab: PresentationContextualRibbonTab,
  mediaKind: PresentationContextualMediaKind | null,
): string {
  switch (tab) {
    case "shape-format":
      return "presentation.home.ribbonShapeFormat";
    case "camera-format":
      return "presentation.cameo.formatTab";
    case "image-format":
      return "presentation.home.ribbonPictureFormat";
    case "graphic-format":
      return "presentation.officeGraphic.formatTab";
    case "chart-design":
      return "presentation.home.ribbonChartDesign";
    case "chart-format":
      return "presentation.ui.chart.formatTab";
    case "media-format":
      return mediaKind === "audio"
        ? "presentation.media.audioFormatTab"
        : "presentation.media.videoFormatTab";
    case "media-playback":
      return "presentation.media.playbackTab";
    case "smartart-design":
      return "presentation.smartartContext.designTab";
    case "smartart-format":
      return "presentation.smartartContext.formatTab";
    case "table-design":
      return "presentation.tableRibbon.designTab";
    case "table-layout":
      return "presentation.tableRibbon.layoutTab";
    case "equation":
      return "equation.ribbon.tab";
    case "record":
      return "presentation.record.tab";
  }
}
