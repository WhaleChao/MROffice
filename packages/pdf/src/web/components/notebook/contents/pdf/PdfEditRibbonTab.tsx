"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type ReactNode, type RefObject, useRef, useState } from "react";

import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { OfficeColorPalettePopover } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import {
  OfficeHomeCommand,
  OfficeHomeMenuTrigger,
  OfficeHomeSplitLayout,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import type { PdfEraseBrushWidth, PdfEraseShape } from "../../../../lib/pdf/page-erase";

import {
  PdfEditRibbonIcon,
  type PdfEditRibbonIconName,
} from "./PdfEditRibbonIcons";
import anatomyStyles from "./PdfInsertRibbonTab.module.css";
import eraseStyles from "./PdfEditEraseMenu.module.css";
import PdfRibbonButton from "./PdfRibbonButton";
import { PdfSurfaceRibbonGroup } from "./PdfSurfaceRibbon";

export interface PdfEditRibbonCommand {
  active?: boolean;
  busy?: boolean;
  controlId?: string;
  disabled?: boolean;
  disabledReason?: string;
  displayLabel?: string;
  icon: PdfEditRibbonIconName;
  label: string;
  menuControlId?: string;
  onInvoke?: (
    trigger: HTMLButtonElement | null,
    resolveTrigger?: () => HTMLButtonElement | null,
  ) => void;
  popup?: "dialog" | "menu";
  testId: string;
  triggerRef?: RefObject<HTMLButtonElement | null>;
}

/**
 * WPS 擦除: the primary face arms the eraser in the chosen colour; its arrow
 * opens the colour picker (自动 / 主题颜色 / 标准色 / 其他颜色…) with the
 * rectangle-or-brush choice and the brush widths below it.
 */
export interface PdfEditEraseCommand extends PdfEditRibbonCommand {
  /** `#rrggbb`, or null for 自动 (the page's own colour under the stroke). */
  color: string | null;
  shape: PdfEraseShape;
  brushWidth: PdfEraseBrushWidth;
  brushWidths: readonly PdfEraseBrushWidth[];
  labels: {
    menu: string;
    automatic: string;
    themeColors: string;
    standardColors: string;
    recentColors: string;
    moreColors: string;
    shape: string;
    rectangle: string;
    brush: string;
    brushWidth: string;
    brushWidthValue: (width: PdfEraseBrushWidth) => string;
    coverNote: string;
  };
  onColorChange: (color: string | null) => void;
  onShapeChange: (shape: PdfEraseShape) => void;
  onBrushWidthChange: (width: PdfEraseBrushWidth) => void;
}

interface PdfEditRibbonTabProps {
  commands: {
    cropPage: PdfEditRibbonCommand;
    deleteFormField: PdfEditRibbonCommand;
    editContent: PdfEditRibbonCommand;
    editFormField: PdfEditRibbonCommand;
    erase: PdfEditEraseCommand;
    insertImage: PdfEditRibbonCommand;
    insertText: PdfEditRibbonCommand;
    /** WPS Edit_EditTextInImage_CurPage (扫描件编辑): recognize this
     * scanned page's text so it can be edited; after Add Image. */
    scanEdit?: PdfEditRibbonCommand;
    linkItems: readonly PdfEditRibbonCommand[];
    links: PdfEditRibbonCommand;
    mergeDocument: PdfEditRibbonCommand;
    shapeItems: readonly PdfEditRibbonCommand[];
    shapes: PdfEditRibbonCommand;
    splitDocument: PdfEditRibbonCommand;
    splitPage: PdfEditRibbonCommand;
    /** WPS Edit's 图章 (Insert_PdfStampTool), after Shapes. */
    stamp?: PdfEditRibbonCommand;
    watermark: PdfEditRibbonCommand;
    watermarkItems: readonly PdfEditRibbonCommand[];
  };
  groupLabels: {
    document: string;
    editContent: string;
    forms: string;
    page: string;
    shapes: string;
    watermark: string;
  };
  pointerTools: ReactNode;
}

function commandDisabled(command: PdfEditRibbonCommand): boolean {
  return Boolean(command.disabled || command.busy || !command.onInvoke);
}

function commandTitle(command: PdfEditRibbonCommand): string {
  if (commandDisabled(command) && command.disabledReason) {
    return `${command.label} — ${command.disabledReason}`;
  }
  return command.label;
}

function EditCommand({
  command,
  order,
}: {
  command: PdfEditRibbonCommand;
  order: string;
}) {
  const internalTriggerRef = useRef<HTMLButtonElement | null>(null);
  const triggerRef = command.triggerRef ?? internalTriggerRef;
  return (
    <PdfRibbonButton
      active={command.active}
      aria-busy={command.busy || undefined}
      aria-haspopup={command.popup}
      buttonRef={triggerRef}
      data-control-id={command.controlId}
      data-disabled-reason={
        commandDisabled(command) ? command.disabledReason : undefined
      }
      data-pdf-edit-order={order}
      disabled={commandDisabled(command)}
      displayLabel={command.displayLabel ?? command.label}
      label={commandTitle(command)}
      onClick={() =>
        command.onInvoke?.(triggerRef.current, () => triggerRef.current)
      }
      testId={command.testId}
    >
      <span className={anatomyStyles.iconSlot} aria-hidden="true">
        <PdfEditRibbonIcon name={command.icon} size={24} />
        {command.busy ? <span className={anatomyStyles.busySpinner} /> : null}
      </span>
    </PdfRibbonButton>
  );
}

function EditMenuItem({
  command,
  menuTriggerRef,
}: {
  command: PdfEditRibbonCommand;
  menuTriggerRef: RefObject<HTMLButtonElement | null>;
}) {
  return (
    <DropdownMenuItem
      aria-label={commandTitle(command)}
      data-testid={command.testId}
      disabled={commandDisabled(command)}
      onSelect={() => command.onInvoke?.(menuTriggerRef.current)}
      title={commandTitle(command)}
    >
      <PdfEditRibbonIcon name={command.icon} size={16} />
      <span>{command.displayLabel ?? command.label}</span>
    </DropdownMenuItem>
  );
}

function EditSplitCommand({
  command,
  menuItems,
  order,
}: {
  command: PdfEditRibbonCommand;
  menuItems: readonly PdfEditRibbonCommand[];
  order: string;
}) {
  const primaryTriggerRef = useRef<HTMLButtonElement | null>(null);
  const menuTriggerRef = useRef<HTMLButtonElement | null>(null);
  const disabled = commandDisabled(command);
  return (
    <DropdownMenu modal={false}>
      <span data-pdf-edit-order={order} style={{ display: "contents" }}>
        <OfficeHomeSplitLayout
          primary={
            <OfficeHomeCommand
              active={command.active}
              armedState={command.busy ? "busy" : undefined}
              buttonRef={primaryTriggerRef}
              controlId={command.controlId}
              disabled={disabled}
              disabledReason={disabled ? command.disabledReason : undefined}
              displayLabel={command.displayLabel ?? command.label}
              icon={
                <span className={anatomyStyles.iconSlot} aria-hidden="true">
                  <PdfEditRibbonIcon name={command.icon} size={24} />
                  {command.busy ? (
                    <span className={anatomyStyles.busySpinner} />
                  ) : null}
                </span>
              }
              label={command.label}
              onClick={() =>
                command.onInvoke?.(primaryTriggerRef.current, () =>
                  primaryTriggerRef.current,
                )
              }
              testId={command.testId}
              variant="large"
            />
          }
          testId={`${command.testId}-split`}
          trigger={
            <DropdownMenuTrigger asChild>
              <OfficeHomeMenuTrigger
                ref={menuTriggerRef}
                className="note-editor-office-toolbar__split-chevron"
                controlId={command.menuControlId}
                disabled={
                  menuItems.length === 0 || menuItems.every(commandDisabled)
                }
                icon={
                  <OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />
                }
                label={command.label}
                testId={`${command.testId}-menu`}
                variant="icon"
              />
            </DropdownMenuTrigger>
          }
        />
      </span>
      <RibbonDropdownMenuContent
        align="start"
        aria-label={command.label}
        className={anatomyStyles.menuContent}
        data-testid={`${command.testId}-menu-content`}
        sideOffset={5}
      >
        {menuItems.map((item) => (
          <EditMenuItem
            key={item.testId}
            command={item}
            menuTriggerRef={menuTriggerRef}
          />
        ))}
      </RibbonDropdownMenuContent>
    </DropdownMenu>
  );
}

function EditEraseOptionButton({
  label,
  onSelect,
  pressed,
  disabled,
  testId,
}: {
  label: string;
  onSelect: () => void;
  pressed: boolean;
  disabled?: boolean;
  testId: string;
}) {
  return (
    <button
      type="button"
      aria-pressed={pressed}
      className={eraseStyles.eraseOption}
      data-testid={testId}
      disabled={disabled}
      onClick={onSelect}
    >
      {label}
    </button>
  );
}

function EditEraseCommand({
  command,
  order,
}: {
  command: PdfEditEraseCommand;
  order: string;
}) {
  const primaryTriggerRef = useRef<HTMLButtonElement | null>(null);
  const menuTriggerRef = useRef<HTMLButtonElement | null>(null);
  const [paletteOpen, setPaletteOpen] = useState(false);
  const disabled = commandDisabled(command);
  const { labels } = command;
  const close = () => setPaletteOpen(false);
  return (
    <span data-pdf-edit-order={order} style={{ display: "contents" }}>
      <OfficeHomeSplitLayout
        primary={
          <OfficeHomeCommand
            active={command.active}
            armedState={command.busy ? "busy" : undefined}
            buttonRef={primaryTriggerRef}
            controlId={command.controlId}
            disabled={disabled}
            disabledReason={disabled ? command.disabledReason : undefined}
            displayLabel={command.displayLabel ?? command.label}
            icon={
              <span
                className={`${anatomyStyles.iconSlot} note-editor-office-toolbar__color-glyph`}
                aria-hidden="true"
              >
                <PdfEditRibbonIcon name={command.icon} size={24} />
                <span
                  className="note-editor-office-toolbar__color-underline"
                  data-testid={`${command.testId}-color`}
                  data-erase-color={command.color ?? "auto"}
                  style={{ backgroundColor: command.color ?? "transparent" }}
                />
              </span>
            }
            label={command.label}
            onClick={() =>
              command.onInvoke?.(primaryTriggerRef.current, () =>
                primaryTriggerRef.current,
              )
            }
            testId={command.testId}
            variant="large"
          />
        }
        testId={`${command.testId}-split`}
        trigger={
          <OfficeHomeMenuTrigger
            ref={menuTriggerRef}
            aria-expanded={paletteOpen}
            aria-haspopup="dialog"
            className="note-editor-office-toolbar__split-chevron"
            controlId={command.menuControlId}
            disabled={disabled}
            disabledReason={disabled ? command.disabledReason : undefined}
            icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
            label={labels.menu}
            onClick={() => setPaletteOpen((open) => !open)}
            testId={`${command.testId}-menu`}
            variant="icon"
          />
        }
      />
      {paletteOpen && !disabled ? (
        <OfficeColorPalettePopover
          anchorRef={menuTriggerRef}
          automaticSelected={command.color === null}
          labels={{
            automatic: labels.automatic,
            themeColors: labels.themeColors,
            standardColors: labels.standardColors,
            recentColors: labels.recentColors,
            moreColors: labels.moreColors,
          }}
          onAutomatic={() => {
            close();
            command.onColorChange(null);
          }}
          onDismiss={close}
          onPick={(color) => {
            close();
            if (color) command.onColorChange(color);
          }}
          testIdPrefix={`${command.testId}-palette`}
          value={command.color ?? undefined}
          footer={
            <div
              className={eraseStyles.eraseOptions}
              data-testid={`${command.testId}-options`}
            >
              <div
                className={eraseStyles.eraseOptionRow}
                role="group"
                aria-label={labels.shape}
              >
                <span className={eraseStyles.eraseOptionLabel}>
                  {labels.shape}
                </span>
                <EditEraseOptionButton
                  label={labels.rectangle}
                  pressed={command.shape === "rect"}
                  onSelect={() => {
                    close();
                    command.onShapeChange("rect");
                  }}
                  testId={`${command.testId}-shape-rect`}
                />
                <EditEraseOptionButton
                  label={labels.brush}
                  pressed={command.shape === "brush"}
                  onSelect={() => {
                    close();
                    command.onShapeChange("brush");
                  }}
                  testId={`${command.testId}-shape-brush`}
                />
              </div>
              <div
                className={eraseStyles.eraseOptionRow}
                role="group"
                aria-label={labels.brushWidth}
              >
                <span className={eraseStyles.eraseOptionLabel}>
                  {labels.brushWidth}
                </span>
                {command.brushWidths.map((width) => (
                  <EditEraseOptionButton
                    key={width}
                    label={labels.brushWidthValue(width)}
                    pressed={command.shape === "brush" && command.brushWidth === width}
                    onSelect={() => {
                      close();
                      command.onBrushWidthChange(width);
                    }}
                    testId={`${command.testId}-width-${width}`}
                  />
                ))}
              </div>
              <p className={eraseStyles.eraseNote}>{labels.coverNote}</p>
            </div>
          }
        />
      ) : null}
    </span>
  );
}

function EditGroup({
  children,
  label,
  testId,
}: {
  children: ReactNode;
  label: string;
  testId: string;
}) {
  return (
    <PdfSurfaceRibbonGroup label={label} tab="edit" testId={testId}>
      {children}
    </PdfSurfaceRibbonGroup>
  );
}

/** PDF-owned adapter over the shared 100px Office Edit surface. */
export default function PdfEditRibbonTab({
  commands,
  groupLabels,
  pointerTools,
}: PdfEditRibbonTabProps) {
  return (
    <>
      {pointerTools}
      {/* WPS CT_Edit RB_Edit is one group: Edit Content, Insert Text, Add
          Image, Scan Edit, Eraser, Path, Stamp, HyperLink. */}
      <EditGroup
        label={groupLabels.editContent}
        testId="pdf-edit-content-group"
      >
        <EditCommand command={commands.editContent} order="02.01" />
        <EditCommand command={commands.insertText} order="03.01" />
        <EditCommand command={commands.insertImage} order="03.02" />
        {commands.scanEdit ? (
          <EditCommand command={commands.scanEdit} order="03.025" />
        ) : null}
        <EditEraseCommand command={commands.erase} order="03.03" />
        <EditSplitCommand
          command={commands.shapes}
          menuItems={commands.shapeItems}
          order="03.04"
        />
        {/* WPS Edit closes the group with 图章 and 超链接 (CT_Edit RB_Edit:
            ... Edit_InsertShapeEx, Insert_PdfStampTool, HyperLink). */}
        {commands.stamp ? <EditCommand command={commands.stamp} order="03.05" /> : null}
        <EditSplitCommand
          command={commands.links}
          menuItems={commands.linkItems}
          order="03.06"
        />
      </EditGroup>
      {/* WPS CT_Edit RB_EditPage: Crop, Split, Watermark. Page numbers and
          headers live on Insert alone (CT_Insertion RB_InsertGraphic). */}
      <EditGroup label={groupLabels.page} testId="pdf-edit-page-group">
        <EditCommand command={commands.cropPage} order="04.01" />
        <EditCommand command={commands.splitPage} order="04.02" />
        <EditSplitCommand
          command={commands.watermark}
          menuItems={commands.watermarkItems}
          order="04.03"
        />
      </EditGroup>
      <EditGroup label={groupLabels.document} testId="pdf-edit-document-group">
        <EditCommand command={commands.mergeDocument} order="06.01" />
        <EditCommand command={commands.splitDocument} order="06.02" />
      </EditGroup>
    </>
  );
}

/** The Edit tab's split command, re-hosted by Shape Editing's Insert Shape. */
export { EditSplitCommand as PdfEditRibbonSplitCommand };
