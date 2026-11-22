/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_WPS_BASE_RIBBON_TABS = ["home", "insert"] as const;

export type PdfWpsBaseRibbonTab = (typeof PDF_WPS_BASE_RIBBON_TABS)[number];

export function isPdfWpsBaseRibbonTab(
  value: string,
): value is PdfWpsBaseRibbonTab {
  return (PDF_WPS_BASE_RIBBON_TABS as readonly string[]).includes(value);
}

export type PdfWpsContextualRibbonTab =
  | "text-format"
  | "picture-format"
  | "comment-tool"
  | "annotation-format"
  | "form-field-format"
  | "link-format"
  /** 表格布局: a table made in MRAI, selected (PdfTableContextRibbon). */
  | "table-format"
  /** 图表设计 and 格式: a chart made in MRAI, selected (PdfChartContextRibbon). */
  | "chart-design"
  | "chart-format";

export type PdfWpsSelectionType =
  | "none"
  | "text"
  | "picture"
  | "text-style"
  | "shape"
  | "path"
  | "measurement"
  | "form-field"
  | "link"
  | "signature"
  | "attachment"
  | "ink"
  | "media"
  | "table"
  | "chart";

export type PdfWpsEditToolMode =
  | "none"
  | "text-edit"
  | "picture-edit"
  | "comment-tool"
  | "form-fill"
  | "form-edit"
  | "graffiti"
  | "path-edit"
  | "measure-tool"
  | "media-edit";

export type PdfWpsDocumentCapability =
  | "standard"
  | "scan"
  | "ebook"
  | "invoice";

export type PdfWpsAuthority = "writable" | "read-only" | "unavailable";
export type PdfWpsWriterSupport = "native" | "preserve-only" | "unsupported";
export type PdfWpsRibbonDensity = "expanded" | "compact";
export type PdfWpsSceneMode = "legacy" | "modern";

export interface PdfWpsRibbonStateInput {
  readonly baseSurface: "pdf";
  readonly requestedBaseTab: PdfWpsBaseRibbonTab;
  readonly selectionType: PdfWpsSelectionType;
  readonly editToolMode: PdfWpsEditToolMode;
  readonly documentCapability: PdfWpsDocumentCapability;
  readonly authority: PdfWpsAuthority;
  readonly ribbonDensity: PdfWpsRibbonDensity;
  readonly writerSupport: PdfWpsWriterSupport;
  readonly sceneMode: PdfWpsSceneMode;
}

export type PdfWpsRibbonContextSource =
  | "selection"
  | "tool-mode"
  | "document-mode"
  | "none";

export interface PdfWpsRibbonStateDecision {
  readonly baseTabs: typeof PDF_WPS_BASE_RIBBON_TABS;
  readonly activeBaseTab: PdfWpsBaseRibbonTab;
  readonly contextualTab: PdfWpsContextualRibbonTab | null;
  readonly contextSource: PdfWpsRibbonContextSource;
  readonly contextEnabled: boolean;
  readonly disabledReason:
    | "read-only"
    | "authority-unavailable"
    | "preserve-only"
    | "writer-unsupported"
    | "document-mode-unavailable"
    | null;
  readonly ribbonDensity: PdfWpsRibbonDensity;
  readonly sceneMode: PdfWpsSceneMode;
}

const ANNOTATION_SELECTIONS = new Set<PdfWpsSelectionType>([
  "text-style",
  "shape",
  "path",
  "measurement",
  "signature",
  "attachment",
  "ink",
  "media",
]);

function selectionContext(
  selectionType: PdfWpsSelectionType,
): PdfWpsContextualRibbonTab | null {
  if (selectionType === "text") return "text-format";
  if (selectionType === "picture") return "picture-format";
  if (selectionType === "form-field") return "form-field-format";
  if (selectionType === "link") return "link-format";
  if (selectionType === "table") return "table-format";
  if (selectionType === "chart") return "chart-design";
  if (ANNOTATION_SELECTIONS.has(selectionType)) return "annotation-format";
  return null;
}

function toolContext(
  editToolMode: PdfWpsEditToolMode,
): PdfWpsContextualRibbonTab | null {
  if (editToolMode === "text-edit") return "text-format";
  if (editToolMode === "picture-edit") return "picture-format";
  if (
    editToolMode === "comment-tool" ||
    editToolMode === "graffiti" ||
    editToolMode === "path-edit" ||
    editToolMode === "measure-tool"
  ) {
    return "comment-tool";
  }
  return null;
}

function toolBaseTab(
  requested: PdfWpsBaseRibbonTab,
  editToolMode: PdfWpsEditToolMode,
): PdfWpsBaseRibbonTab {
  if (
    editToolMode === "comment-tool" ||
    editToolMode === "graffiti" ||
    editToolMode === "path-edit" ||
    editToolMode === "measure-tool"
  ) {
    return "annotate";
  }
  if (
    editToolMode === "form-fill" ||
    editToolMode === "form-edit" ||
    editToolMode === "media-edit" ||
    editToolMode === "picture-edit"
  ) {
    return "edit";
  }
  return requested;
}

export function resolvePdfWpsRibbonState(
  input: PdfWpsRibbonStateInput,
): PdfWpsRibbonStateDecision {
  const activeBaseTab = toolBaseTab(input.requestedBaseTab, input.editToolMode);
  const fromSelection = selectionContext(input.selectionType);
  const fromTool = fromSelection ? null : toolContext(input.editToolMode);
  const contextualTab = fromSelection ?? fromTool;
  const contextSource: PdfWpsRibbonContextSource = fromSelection
    ? "selection"
    : fromTool
      ? "tool-mode"
      : input.documentCapability !== "standard"
        ? "document-mode"
        : "none";

  if (!contextualTab) {
    return {
      baseTabs: PDF_WPS_BASE_RIBBON_TABS,
      activeBaseTab,
      contextualTab: null,
      contextSource,
      contextEnabled: false,
      disabledReason:
        contextSource === "document-mode" ? "document-mode-unavailable" : null,
      ribbonDensity: input.ribbonDensity,
      sceneMode: input.sceneMode,
    };
  }
  if (input.writerSupport === "unsupported") {
    return {
      baseTabs: PDF_WPS_BASE_RIBBON_TABS,
      activeBaseTab,
      contextualTab: null,
      contextSource,
      contextEnabled: false,
      disabledReason: "writer-unsupported",
      ribbonDensity: input.ribbonDensity,
      sceneMode: input.sceneMode,
    };
  }

  const disabledReason =
    input.authority === "read-only"
      ? "read-only"
      : input.authority === "unavailable"
        ? "authority-unavailable"
        : input.writerSupport === "preserve-only"
          ? "preserve-only"
          : null;
  return {
    baseTabs: PDF_WPS_BASE_RIBBON_TABS,
    activeBaseTab,
    contextualTab,
    contextSource,
    contextEnabled: disabledReason === null,
    disabledReason,
    ribbonDensity: input.ribbonDensity,
    sceneMode: input.sceneMode,
  };
}
