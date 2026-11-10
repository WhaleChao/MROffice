"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Dispatch, SetStateAction, RefObject } from "react";
import {
  PdfSurfaceRibbonGroup,
  PdfSurfaceRibbonCommand,
} from "../PdfSurfaceRibbon";
import {
  PdfInsertRibbonIcon,
} from "../PdfInsertRibbonIcons";
import type { CropBoxEntry } from "../../../../../lib/pdf/crop-box";
import type { PdfEditorRibbonTab } from "./editorRibbonTabs";
import type {
  PdfInsertRibbonCommand,
} from "../PdfInsertRibbonTab";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";

export interface PdfWindowPageTabGroupsProps {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  cropBoxes: CropBoxEntry[];
  cropOpen: boolean;
  setCropOpen: Dispatch<SetStateAction<boolean>>;
  cropToggleRef: RefObject<HTMLButtonElement | null>;
  cropDialogReturnFocusTargetRef: RefObject<() => HTMLElement | null>;
  editorRibbonTab: PdfEditorRibbonTab;
  currentPage: number;
  pageStructureRibbonActions: PdfInsertRibbonCommand[];
}

export function PdfWindowPageTabGroups({
  t,
  cropBoxes,
  cropOpen,
  setCropOpen,
  cropToggleRef,
  cropDialogReturnFocusTargetRef,
  editorRibbonTab,
  currentPage,
  pageStructureRibbonActions,
}: PdfWindowPageTabGroupsProps) {
  return (
    editorRibbonTab === "page" ? (
      <PdfSurfaceRibbonGroup
        label={t("pdf.pageWindow.pageStructure")}
        tab="page"
        testId="pdf-editor-page-surface"
      >
        {pageStructureRibbonActions.map((command) => (
          <PdfSurfaceRibbonCommand
            key={command.testId}
            active={command.active}
            ariaExpanded={command.expanded}
            ariaHasPopup={command.popup}
            busy={command.busy}
            controlId={command.controlId}
            disabled={Boolean(command.disabled || !command.onInvoke)}
            disabledReason={command.disabledReason}
            displayLabel={command.displayLabel}
            icon={<PdfInsertRibbonIcon name={command.icon} size={24} />}
            label={command.label}
            onClick={(trigger) => command.onInvoke?.(trigger)}
            testId={command.testId}
          />
        ))}
        <PdfSurfaceRibbonCommand
          active={
            cropOpen ||
            cropBoxes.some((entry) => entry.page === currentPage)
          }
          buttonRef={cropToggleRef}
          controlId="PDFCTL.SURFACE.CROP_PAGE"
          icon={<PdfInsertRibbonIcon name="cropPage" size={24} />}
          label={t("pdf.crop.button")}
          onClick={() => {
            cropDialogReturnFocusTargetRef.current = () =>
              cropToggleRef.current;
            setCropOpen((value) => !value);
          }}
          testId="pdf-crop-toggle"
          toggle
        />
      </PdfSurfaceRibbonGroup>
    ) : null
  );
}
