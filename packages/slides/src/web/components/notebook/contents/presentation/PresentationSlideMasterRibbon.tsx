"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Slide Master tab (PPT-B030): the tab PowerPoint shows, first and
 * selected, while Slide Master view is open, with the desktop's groups in its
 * order -- Edit Master (Insert Slide Master, Insert Layout, Delete, Rename,
 * Preserve), Master Layout (Master Layout, Insert Placeholder, Title,
 * Footers), Edit Theme (Themes), Background (Colors, Fonts, Effects,
 * Background Styles, Hide Background Graphics), Size (Slide Size) and Close
 * (Close Master View).
 *
 * The window passes the galleries it already owns (Themes, Colors, Fonts,
 * Background Styles, Slide Size) as menus, so the Design tab and this tab run
 * one command each. A command this editor cannot carry out on the selected
 * master yet is shown disabled with the reason, never as a control that
 * pretends to work. Delete addresses the selected card: a layout no slide
 * uses, or a master no slide uses while another master remains.
 */

import type { ReactNode } from "react";

import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  OFFICE_PLACEHOLDER_KINDS,
  type OfficePlaceholderKind,
} from "@mrnote/core/web/lib/presentation-office-master-model.ts";

import {
  OfficeHomeCheckbox,
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeRibbonGroup,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import {
  FolderPlus,
  LayoutTemplate,
  Pen,
  Plus,
  Square,
  Trash2,
  X,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";

export interface PresentationSlideMasterRibbonLabels {
  readonly editMaster: string;
  readonly insertSlideMaster: string;
  readonly insertLayout: string;
  readonly delete: string;
  readonly rename: string;
  readonly preserve: string;
  readonly masterLayoutGroup: string;
  readonly masterLayout: string;
  readonly insertPlaceholder: string;
  readonly title: string;
  readonly footers: string;
  readonly editTheme: string;
  readonly background: string;
  readonly hideBackgroundGraphics: string;
  readonly size: string;
  readonly close: string;
  readonly closeMasterView: string;
}

export interface PresentationSlideMasterRibbonProps {
  readonly labels: PresentationSlideMasterRibbonLabels;
  /**
   * Why the master's structure (Edit Master, Master Layout, Hide Background
   * Graphics) cannot be edited here yet; those commands are disabled with it.
   */
  readonly structureUnavailableReason: string;
  /** Whether the selected layout shows a title / the footers (read-only state). */
  readonly layoutHasTitle: boolean;
  readonly layoutHasFooters: boolean;
  readonly layoutHidesBackgroundGraphics: boolean;
  /** Design ▸ Themes, Variants ▸ Colors / Fonts / Effects / Background Styles, Slide Size. */
  readonly themesMenu: ReactNode;
  readonly colorsMenu: ReactNode;
  readonly fontsMenu: ReactNode;
  readonly effectsMenu: ReactNode;
  readonly backgroundStylesMenu: ReactNode;
  readonly slideSizeMenu: ReactNode;
  readonly onClose: () => void;
  /** Rename the selected master or layout; absent where it cannot be renamed. */
  readonly onRename?: () => void;
  /** Preserve on the master: its state and toggle; absent where unavailable. */
  readonly preserve?: { readonly pressed: boolean; readonly onToggle: () => void };
  /**
   * Delete on the selected layout (PPT-B030): absent where it cannot go,
   * with `deleteUnavailableReason` saying why (in use, the last layout).
   */
  readonly onDelete?: () => void;
  readonly deleteUnavailableReason?: string;
  /** Insert Placeholder on the selected layout; absent on the master card. */
  readonly onInsertPlaceholder?: (kind: OfficePlaceholderKind) => void;
  readonly placeholderKindLabels?: Readonly<Record<OfficePlaceholderKind, string>>;
  /** Title / Footers on the selected layout; absent where they cannot change. */
  readonly onTitleChange?: (on: boolean) => void;
  readonly onFootersChange?: (on: boolean) => void;
  /** Why Master Layout's commands are unavailable on the selected card. */
  readonly layoutCommandsUnavailableReason?: string;
  /** Edit Master ▸ Insert Layout (PPT-B030); absent where it cannot run. */
  readonly onInsertLayout?: () => void;
  /** Edit Master ▸ Insert Slide Master (PPT-B030); absent where it cannot run. */
  readonly onInsertSlideMaster?: () => void;
  /** Master Layout ▸ Master Layout, the dialog (on the slide master card only). */
  readonly onMasterLayout?: () => void;
  readonly masterLayoutUnavailableReason?: string;
  /** Background ▸ Hide Background Graphics on the selected layout. */
  readonly onHideBackgroundGraphicsChange?: (hide: boolean) => void;
  readonly hideBackgroundGraphicsUnavailableReason?: string;
}

export function PresentationSlideMasterRibbon({
  labels,
  structureUnavailableReason,
  layoutHasTitle,
  layoutHasFooters,
  layoutHidesBackgroundGraphics,
  themesMenu,
  colorsMenu,
  fontsMenu,
  effectsMenu,
  backgroundStylesMenu,
  slideSizeMenu,
  onClose,
  onRename,
  preserve,
  onDelete,
  deleteUnavailableReason,
  onInsertPlaceholder,
  placeholderKindLabels,
  onTitleChange,
  onFootersChange,
  layoutCommandsUnavailableReason,
  onInsertLayout,
  onInsertSlideMaster,
  onMasterLayout,
  masterLayoutUnavailableReason,
  onHideBackgroundGraphicsChange,
  hideBackgroundGraphicsUnavailableReason,
}: PresentationSlideMasterRibbonProps) {
  const layoutReason = layoutCommandsUnavailableReason ?? structureUnavailableReason;
  const unavailable = {
    disabled: true,
    disabledReason: structureUnavailableReason,
    onClick: () => {},
  } as const;
  return (
    <>
      <OfficeHomeRibbonGroup
        label={labels.editMaster}
        ribbon="slide-master"
        testId="presentation-slide-master-edit-group"
      >
        <OfficeHomeCommandRow>
          <OfficeHomeCommand
            {...(onInsertSlideMaster ? { onClick: onInsertSlideMaster } : unavailable)}
            icon={<FolderPlus size={24} />}
            label={labels.insertSlideMaster}
            testId="presentation-slide-master-insert-master"
            variant="large"
          />
          <OfficeHomeCommand
            {...(onInsertLayout ? { onClick: onInsertLayout } : unavailable)}
            icon={<Plus size={24} />}
            label={labels.insertLayout}
            testId="presentation-slide-master-insert-layout"
            variant="large"
          />
          <OfficeHomeCommandStack>
            <OfficeHomeCommand
              {...(onDelete
                ? { onClick: onDelete }
                : { ...unavailable, disabledReason: deleteUnavailableReason ?? structureUnavailableReason })}
              icon={<Trash2 size={16} />}
              label={labels.delete}
              testId="presentation-slide-master-delete"
            />
            <OfficeHomeCommand
              {...(onRename ? { onClick: onRename } : unavailable)}
              icon={<Pen size={16} />}
              label={labels.rename}
              testId="presentation-slide-master-rename"
            />
            <OfficeHomeCommand
              {...(preserve ? { onClick: preserve.onToggle, active: preserve.pressed } : unavailable)}
              icon={<Square size={16} />}
              label={labels.preserve}
              testId="presentation-slide-master-preserve"
            />
          </OfficeHomeCommandStack>
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup
        label={labels.masterLayoutGroup}
        ribbon="slide-master"
        testId="presentation-slide-master-layout-group"
      >
        <OfficeHomeCommandRow>
          <OfficeHomeCommand
            {...(onMasterLayout
              ? { onClick: onMasterLayout }
              : {
                  ...unavailable,
                  disabledReason: masterLayoutUnavailableReason ?? structureUnavailableReason,
                })}
            icon={<LayoutTemplate size={24} />}
            label={labels.masterLayout}
            testId="presentation-slide-master-master-layout"
            variant="large"
          />
          {onInsertPlaceholder && placeholderKindLabels ? (
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <OfficeHomeMenuTrigger
                  hasMenuIndicator
                  icon={<Plus size={24} />}
                  label={labels.insertPlaceholder}
                  testId="presentation-slide-master-insert-placeholder"
                  variant="large"
                />
              </DropdownMenuTrigger>
              <RibbonDropdownMenuContent align="start" data-testid="presentation-slide-master-insert-placeholder-menu">
                {OFFICE_PLACEHOLDER_KINDS.map((kind) => (
                  <DropdownMenuItem
                    key={kind}
                    onSelect={() => onInsertPlaceholder(kind)}
                    data-testid={`presentation-slide-master-insert-placeholder-${kind}`}
                  >
                    {placeholderKindLabels[kind]}
                  </DropdownMenuItem>
                ))}
              </RibbonDropdownMenuContent>
            </DropdownMenu>
          ) : (
            <OfficeHomeCommand
              {...unavailable}
              disabledReason={layoutReason}
              icon={<Plus size={24} />}
              label={labels.insertPlaceholder}
              testId="presentation-slide-master-insert-placeholder"
              variant="large"
            />
          )}
          <OfficeHomeCommandStack>
            <OfficeHomeCheckbox
              checked={layoutHasTitle}
              disabled={!onTitleChange}
              title={onTitleChange ? undefined : layoutReason}
              label={labels.title}
              onCheckedChange={(checked) => onTitleChange?.(checked)}
              testId="presentation-slide-master-title"
            />
            <OfficeHomeCheckbox
              checked={layoutHasFooters}
              disabled={!onFootersChange}
              title={onFootersChange ? undefined : layoutReason}
              label={labels.footers}
              onCheckedChange={(checked) => onFootersChange?.(checked)}
              testId="presentation-slide-master-footers"
            />
          </OfficeHomeCommandStack>
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup
        label={labels.editTheme}
        ribbon="slide-master"
        testId="presentation-slide-master-theme-group"
      >
        <OfficeHomeCommandRow>{themesMenu}</OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup
        label={labels.background}
        ribbon="slide-master"
        testId="presentation-slide-master-background-group"
      >
        <OfficeHomeCommandRow>
          <OfficeHomeCommandStack>
            {colorsMenu}
            {fontsMenu}
            {effectsMenu}
          </OfficeHomeCommandStack>
          <OfficeHomeCommandStack>
            {backgroundStylesMenu}
            <OfficeHomeCheckbox
              checked={layoutHidesBackgroundGraphics}
              disabled={!onHideBackgroundGraphicsChange}
              title={
                onHideBackgroundGraphicsChange
                  ? undefined
                  : (hideBackgroundGraphicsUnavailableReason ?? structureUnavailableReason)
              }
              label={labels.hideBackgroundGraphics}
              onCheckedChange={(checked) => onHideBackgroundGraphicsChange?.(checked)}
              testId="presentation-slide-master-hide-background-graphics"
            />
          </OfficeHomeCommandStack>
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup
        label={labels.size}
        ribbon="slide-master"
        testId="presentation-slide-master-size-group"
      >
        <OfficeHomeCommandRow>{slideSizeMenu}</OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup
        label={labels.close}
        ribbon="slide-master"
        testId="presentation-slide-master-close-group"
      >
        <OfficeHomeCommandRow>
          <OfficeHomeCommand
            controlId="PPTCTL.RIBBON.SLIDE_MASTER.CLOSE"
            icon={<X size={24} />}
            label={labels.closeMasterView}
            onClick={onClose}
            testId="presentation-slide-master-close"
            variant="large"
          />
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
    </>
  );
}
