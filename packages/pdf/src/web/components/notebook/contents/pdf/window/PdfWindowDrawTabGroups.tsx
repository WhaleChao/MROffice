"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Dispatch, SetStateAction } from "react";

import { PdfDrawRibbonTab, type PdfDrawRibbonModel, type PdfInkTool } from "../PdfDrawRibbonTab";
import type { PdfEditorRibbonTab } from "./editorRibbonTabs";
import type { PdfInkPenCommands } from "./usePdfInkPens";
import type { useAutosavedPageDocument } from "@mrnote/core/web/hooks/useAutosavedPageDocument.ts";
import type { DocumentInkEraser } from "@mrnote/core/web/lib/document-ink.ts";
import type { PdfInkStroke } from "../PdfInkOverlay";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";

export interface PdfWindowDrawTabGroupsProps {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  workspaceWritable: boolean;
  hasWorkspaceWriteAuthority: () => boolean;
  editorRibbonTab: PdfEditorRibbonTab;
  autosave: ReturnType<typeof useAutosavedPageDocument>;
  workspaceWriteDeniedMessage: string;
  inkTool: PdfInkTool | null;
  setInkTool: (tool: PdfInkTool | null) => void;
  inkEraserKind: DocumentInkEraser;
  setInkEraserKind: Dispatch<SetStateAction<DocumentInkEraser>>;
  inkPenCommands: PdfInkPenCommands;
  managedInkStrokes: PdfInkStroke[];
  undoLastInkStroke: () => void;
  trackpad: PdfDrawRibbonModel["trackpad"];
  hideRibbon: (() => void) | null;
}






export function PdfWindowDrawTabGroups({
  t,
  workspaceWritable,
  hasWorkspaceWriteAuthority,
  editorRibbonTab,
  autosave,
  workspaceWriteDeniedMessage,
  inkTool,
  setInkTool,
  inkEraserKind,
  setInkEraserKind,
  inkPenCommands,
  managedInkStrokes,
  undoLastInkStroke,
  trackpad,
  hideRibbon,
}: PdfWindowDrawTabGroupsProps) {
  if (editorRibbonTab !== "draw") return null;
  const writable = workspaceWritable && !autosave.readOnly;
  const model: PdfDrawRibbonModel = {
    writable,
    disabledReason: writable ? undefined : workspaceWriteDeniedMessage,
    tool: inkTool,
    setTool: (tool) => {
      if (tool !== null && !hasWorkspaceWriteAuthority()) return;
      setInkTool(tool);
    },
    eraserKind: inkEraserKind,
    chooseEraserKind: (kind) => {
      if (!hasWorkspaceWriteAuthority()) return;
      setInkEraserKind(kind);
      setInkTool("eraser");
    },
    inkPresent: managedInkStrokes.length > 0,
    pens: inkPenCommands,
    undoLastStroke: {
      run: undoLastInkStroke,
      disabled: managedInkStrokes.length === 0,
    },
    trackpad,
    hideRibbon,
  };
  return <PdfDrawRibbonTab t={t} model={model} />;
}
