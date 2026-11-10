"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  ChevronsRight,
  CircleDot,
  Download,
  Ellipsis,
  ExternalLink,
  FileCode2,
  FlipVertical2,
  FormInput,
  Hash,
  History,
  Info,
  ImageIcon,
  Images,
  Languages,
  ListChecks,
  ListFilter,
  ListTree,
  MonitorUp,
  PencilRuler,
  Printer,
  RotateCcw,
  Save,
  ScanText,
  Search,
  Signature,
  Square,
  Trash2,
  Upload,
  X,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useRef, type ReactNode, type RefObject } from "react";

import { OfficeHomeCommand, OfficeHomeCommandStack, OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";

import {
  PdfHomeRibbonIcon,
  type PdfHomeRibbonIconName,
} from "./PdfHomeRibbonIcons";
import PdfRibbonButton from "./PdfRibbonButton";
import { PdfSurfaceRibbonGroup } from "./PdfSurfaceRibbon";
import ribbonLayout from "./pdfRibbonLayout.module.css";

export interface PdfFeaturedCommand {
  active?: boolean;
  buttonRef?: RefObject<HTMLButtonElement | null>;
  disabled?: boolean;
  disabledReason?: string;
  displayLabel?: string;
  /**
   * Withhold the command entirely. Reserved for a host that can never satisfy
   * this command for the life of the viewer — a preview with no data item can
   * never authenticate a source, so the output commands gated on one would sit
   * grey forever and read as "not yet". `disabled` stays for the conditions a
   * user can actually clear.
   */
  hidden?: boolean;
  icon?:
    | PdfHomeRibbonIconName
    | "ocr"
    | "deskew"
    | "translation"
    | "screenRecording"
    | "imagesToPdf"
    | "markdownToPdf"
    | "cadToPdf"
    | "stop"
    | "formFields"
    | "deleteField"
    | "checkbox"
    | "radio"
    | "signatureField"
    | "printButton"
    | "navigationButton"
    | "viewButton"
    | "saveButton"
    | "goToPageButton"
    | "webLinkButton"
    | "dropdown"
    | "listbox"
    | "imageField"
    | "versions"
    | "properties"
    | "importData"
    | "exportData"
    | "resetData"
    | "cancel"
    | "more";
  label: string;
  onInvoke?: (trigger: HTMLButtonElement | null) => void;
  testId: string;
}

export interface PdfFeaturedRibbonTabProps {
  visible: boolean;
  pointerTools: ReactNode;
  status?: string | null;
  groupLabels: {
    convert: string;
    extract: string;
    service: string;
  };
  commands: {
    word: PdfFeaturedCommand;
    excel: PdfFeaturedCommand;
    ppt: PdfFeaturedCommand;
    txt: PdfFeaturedCommand;
    cad: PdfFeaturedCommand;
    imageOnly: PdfFeaturedCommand;
    exportImages: PdfFeaturedCommand;
    imagesToPdf: PdfFeaturedCommand;
    /** WPS' Markdown to PDF; optional for hosts with no server converter. */
    markdownToPdf?: PdfFeaturedCommand;
    /** WPS' CAD to PDF (DXF drawings); optional like Markdown to PDF. */
    cadToPdf?: PdfFeaturedCommand;
    splitFile: PdfFeaturedCommand;
    merge: PdfFeaturedCommand;
    extractText: PdfFeaturedCommand;
    extractImages: PdfFeaturedCommand;
    recognizeText: PdfFeaturedCommand;
    screenshotText: PdfFeaturedCommand;
    compress: PdfFeaturedCommand;
    translate: PdfFeaturedCommand;
    readAloud: PdfFeaturedCommand;
    stopReadAloud?: PdfFeaturedCommand;
    screenRecording: PdfFeaturedCommand;
    stopScreenRecording?: PdfFeaturedCommand;
  };
  statusTestId?: string;
  /** The form, data and document tools the retired "more" menu held. */
  globalTools?: PdfGlobalToolsOverflowProps;
}

/** An overflow entry that opens a menu of commands instead of running one. */
export interface PdfFeaturedMenuCommand {
  command: PdfFeaturedCommand;
  items: readonly PdfFeaturedCommand[];
}

export interface PdfGlobalToolsOverflowProps {
  label: string;
  commands: {
    fillForm: PdfFeaturedCommand;
    /** 以角色填写 (ONLYOFFICE's View form as a role): offered when the form
     * names roles -- anyone, or one role, whose fields alone take input. */
    fillFormAs?: PdfFeaturedMenuCommand;
    /** New fields: the field kinds the authoring card can create. */
    addFormField: PdfFeaturedMenuCommand;
    /** New action buttons: reset, print, navigation and the other kinds. */
    addFormAction: PdfFeaturedMenuCommand;
    editFormField: PdfFeaturedCommand;
    deleteFormField: PdfFeaturedCommand;
    detectFormFields: PdfFeaturedCommand;
    cancelDetectedFields?: PdfFeaturedCommand;
    exportFormData: PdfFeaturedCommand;
    importFormData: PdfFeaturedCommand;
    resetForm: PdfFeaturedCommand;
    versions: PdfFeaturedCommand;
    properties: PdfFeaturedCommand;
    deskew: PdfFeaturedCommand;
    moreTools: PdfFeaturedCommand;
  };
}

/**
 * A command with no `onInvoke` has nothing to run, so it is disabled for the
 * same reason a `disabled` one is. Every other PDF ribbon tab folds that into
 * its own disabled predicate (`PdfFileRibbonTab`, `PdfProtectRibbonTab`,
 * `PdfCommentRibbonTab`, `PdfInsertRibbonTab`, `PdfEditRibbonTab`,
 * `PdfHomeCompactRibbon`); this tab used to trust its single caller to pair
 * the two by hand, which makes the next caller's omission a live button that
 * does nothing.
 */
function commandDisabled(command: PdfFeaturedCommand): boolean {
  return Boolean(command.disabled || !command.onInvoke);
}

function commandLabel(command: PdfFeaturedCommand): string {
  return commandDisabled(command) && command.disabledReason
    ? `${command.label}: ${command.disabledReason}`
    : command.label;
}

function icon(command: PdfFeaturedCommand, size: number) {
  switch (command.icon) {
    case "ocr":
      return <ScanText size={size} />;
    case "deskew":
      return <FlipVertical2 size={size} />;
    case "translation":
      return <Languages size={size} />;
    case "screenRecording":
      return <MonitorUp size={size} />;
    case "imagesToPdf":
      return <Images size={size} />;
    case "markdownToPdf":
      return <FileCode2 size={size} />;
    case "cadToPdf":
      return <PencilRuler size={size} />;
    case "stop":
      return <Square size={size} />;
    case "formFields":
      return <FormInput size={size} />;
    case "deleteField":
      return <Trash2 size={size} />;
    case "checkbox":
      return <ListChecks size={size} />;
    case "radio":
      return <CircleDot size={size} />;
    case "signatureField":
      return <Signature size={size} />;
    case "printButton":
      return <Printer size={size} />;
    case "navigationButton":
      return <ChevronsRight size={size} />;
    case "viewButton":
      return <Search size={size} />;
    case "saveButton":
      return <Save size={size} />;
    case "goToPageButton":
      return <Hash size={size} />;
    case "webLinkButton":
      return <ExternalLink size={size} />;
    case "dropdown":
      return <ListFilter size={size} />;
    case "listbox":
      return <ListTree size={size} />;
    case "imageField":
      return <ImageIcon size={size} />;
    case "versions":
      return <History size={size} />;
    case "properties":
      return <Info size={size} />;
    case "importData":
      return <Upload size={size} />;
    case "exportData":
      return <Download size={size} />;
    case "resetData":
      return <RotateCcw size={size} />;
    case "cancel":
      return <X size={size} />;
    case "more":
      return <Ellipsis size={size} />;
    default:
      return command.icon ? (
        <PdfHomeRibbonIcon name={command.icon} size={size} />
      ) : (
        <Ellipsis size={size} />
      );
  }
}

function Command({
  command,
  variant = "large",
}: {
  command: PdfFeaturedCommand;
  variant?: "large" | "compact";
}) {
  const fallbackTriggerRef = useRef<HTMLButtonElement>(null);
  const triggerRef = command.buttonRef ?? fallbackTriggerRef;
  if (command.hidden) return null;
  if (variant === "large") {
    return (
      <PdfRibbonButton
        active={command.active}
        buttonRef={triggerRef}
        data-disabled-reason={
          commandDisabled(command) ? command.disabledReason : undefined
        }
        disabled={commandDisabled(command)}
        displayLabel={command.displayLabel ?? command.label}
        label={commandLabel(command)}
        onClick={() => command.onInvoke?.(triggerRef.current)}
        testId={command.testId}
      >
        <span aria-hidden="true">{icon(command, 24)}</span>
      </PdfRibbonButton>
    );
  }
  return (
    <OfficeHomeCommand
      active={command.active}
      buttonRef={triggerRef}
      disabled={commandDisabled(command)}
      icon={icon(command, 20)}
      label={commandLabel(command)}
      displayLabel={command.displayLabel ?? command.label}
      onClick={() => command.onInvoke?.(triggerRef.current)}
      testId={command.testId}
      variant={variant}
    />
  );
}

function FeaturedGroup({
  children,
  label,
  testId,
}: {
  children: ReactNode;
  label: string;
  testId: string;
}) {
  return (
    <PdfSurfaceRibbonGroup label={label} tab="tools" testId={testId}>
      {children}
    </PdfSurfaceRibbonGroup>
  );
}

function OverflowCommand({ command }: { command: PdfFeaturedCommand }) {
  const fallbackTriggerRef = useRef<HTMLButtonElement>(null);
  const triggerRef = command.buttonRef ?? fallbackTriggerRef;
  const accessibleLabel = commandLabel(command);
  if (command.hidden) return null;
  return (
    <PdfRibbonButton
      active={command.active}
      big={false}
      buttonRef={triggerRef}
      captioned
      data-disabled-reason={
        commandDisabled(command) ? command.disabledReason : undefined
      }
      disabled={commandDisabled(command)}
      displayLabel={command.displayLabel ?? command.label}
      label={accessibleLabel}
      onClick={() => command.onInvoke?.(triggerRef.current)}
      testId={`${command.testId}-overflow-item`}
    >
      {icon(command, 20)}
    </PdfRibbonButton>
  );
}

 
function stacked(entries: readonly (ReactNode | null)[]): ReactNode {
  const shown = entries.filter((entry): entry is ReactNode => entry !== null);
  const lanes: ReactNode[] = [];
  for (let start = 0; start < shown.length; start += 3) {
    lanes.push(
      <OfficeHomeCommandStack key={start} rows={3} className={ribbonLayout.keepsWidth}>
        {shown.slice(start, start + 3)}
      </OfficeHomeCommandStack>,
    );
  }
  return lanes;
}

function OverflowMenuCommand({ command, items }: PdfFeaturedMenuCommand) {
  const fallbackTriggerRef = useRef<HTMLButtonElement>(null);
  const triggerRef = command.buttonRef ?? fallbackTriggerRef;
  const disabled = Boolean(
    command.disabled || items.length === 0 || items.every(commandDisabled),
  );
  const accessibleLabel =
    disabled && command.disabledReason
      ? `${command.label} — ${command.disabledReason}`
      : command.label;
  if (command.hidden) return null;
  return (
    <DropdownMenu modal={false}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          ref={triggerRef}
          disabled={disabled}
          disabledReason={disabled ? command.disabledReason : undefined}
          displayLabel={command.displayLabel ?? command.label}
          hasMenuIndicator
          icon={icon(command, 20)}
          label={command.label}
          testId={`${command.testId}-overflow-item`}
          title={accessibleLabel}
          variant="compact"
        />
      </DropdownMenuTrigger>
      <RibbonDropdownMenuContent
        align="start"
        aria-label={command.label}
        data-testid={`${command.testId}-menu-content`}
        sideOffset={5}
      >
        {items.map((item) => (
          <DropdownMenuItem
            key={item.testId}
            aria-label={commandLabel(item)}
            disabled={commandDisabled(item)}
            title={commandLabel(item)}
            data-testid={item.testId}
            onSelect={() => item.onInvoke?.(triggerRef.current)}
          >
            {icon(item, 16)}
            <span>{item.displayLabel ?? item.label}</span>
          </DropdownMenuItem>
        ))}
      </RibbonDropdownMenuContent>
    </DropdownMenu>
  );
}

export function PdfGlobalToolsOverflow({
  label,
  commands,
}: PdfGlobalToolsOverflowProps) {
  return (
    <span
      className="contents"
      role="group"
      aria-label={label}
      data-testid="pdf-global-tools-overflow"
    >
      <span
        className="contents"
        role="group"
        aria-label={commands.fillForm.label}
        data-testid="pdf-global-tools-form-group"
      >
        {stacked([
          commands.fillForm.hidden ? null : <OverflowCommand key="fill" command={commands.fillForm} />,
          commands.fillFormAs && !commands.fillFormAs.command.hidden
            ? <OverflowMenuCommand key="fill-as" {...commands.fillFormAs} />
            : null,
          commands.addFormField.command.hidden ? null : <OverflowMenuCommand key="add-field" {...commands.addFormField} />,
          commands.addFormAction.command.hidden ? null : <OverflowMenuCommand key="add-action" {...commands.addFormAction} />,
          commands.editFormField.hidden ? null : <OverflowCommand key="edit" command={commands.editFormField} />,
          commands.deleteFormField.hidden ? null : <OverflowCommand key="delete" command={commands.deleteFormField} />,
          commands.detectFormFields.hidden ? null : <OverflowCommand key="detect" command={commands.detectFormFields} />,
          commands.cancelDetectedFields && !commands.cancelDetectedFields.hidden
            ? <OverflowCommand key="cancel-detect" command={commands.cancelDetectedFields} />
            : null,
        ])}
      </span>
      <span
        className="contents"
        role="group"
        aria-label={commands.exportFormData.label}
        data-testid="pdf-global-tools-data-group"
      >
        {stacked([
          commands.exportFormData.hidden ? null : <OverflowCommand key="export" command={commands.exportFormData} />,
          commands.importFormData.hidden ? null : <OverflowCommand key="import" command={commands.importFormData} />,
          commands.resetForm.hidden ? null : <OverflowCommand key="reset" command={commands.resetForm} />,
        ])}
      </span>
      <span
        className="contents"
        role="group"
        aria-label={label}
        data-testid="pdf-global-tools-document-group"
      >
        {stacked([
          commands.versions.hidden ? null : <OverflowCommand key="versions" command={commands.versions} />,
          commands.properties.hidden ? null : <OverflowCommand key="properties" command={commands.properties} />,
          commands.deskew.hidden ? null : <OverflowCommand key="deskew" command={commands.deskew} />,
          commands.moreTools.hidden ? null : <OverflowCommand key="more" command={commands.moreTools} />,
        ])}
      </span>
    </span>
  );
}

export function PdfFeaturedRibbonTab({
  visible,
  status,
  statusTestId,
  groupLabels,
  commands,
  pointerTools,
  globalTools,
}: PdfFeaturedRibbonTabProps) {
  if (!visible) return null;
  return (
    <>
      {pointerTools}
      <FeaturedGroup
        label={groupLabels.convert}
        testId="pdf-featured-convert-group"
      >
        <Command command={commands.word} />
        <Command command={commands.excel} />
        <Command command={commands.ppt} />
        <Command command={commands.txt} />
        <Command command={commands.cad} />
        <Command command={commands.imageOnly} />
        {/* WPS keeps the file outputs in the conversion run, then a
            separator before the extract-and-OCR commands. */}
        <Command command={commands.exportImages} />
        <Command command={commands.imagesToPdf} />
        {/* WPS: Photo2PDF, then CAD2PDF, then the Markdown converter. */}
        {commands.cadToPdf ? <Command command={commands.cadToPdf} /> : null}
        {commands.markdownToPdf ? (
          <Command command={commands.markdownToPdf} />
        ) : null}
        <Command command={commands.splitFile} />
        <Command command={commands.merge} />
      </FeaturedGroup>
      <FeaturedGroup
        label={groupLabels.extract}
        testId="pdf-featured-extract-group"
      >
        <Command command={commands.extractText} />
        <Command command={commands.recognizeText} />
        <Command command={commands.extractImages} />
        <Command command={commands.screenshotText} />
      </FeaturedGroup>
      <FeaturedGroup
        label={groupLabels.service}
        testId="pdf-featured-service-group"
      >
        <Command command={commands.compress} />
        <Command command={commands.translate} />
        <Command command={commands.readAloud} />
        {commands.stopReadAloud ? (
          <Command command={commands.stopReadAloud} variant="compact" />
        ) : null}
        <Command command={commands.screenRecording} />
        {commands.stopScreenRecording ? (
          <Command command={commands.stopScreenRecording} variant="compact" />
        ) : null}
      </FeaturedGroup>
      {globalTools ? (
        <FeaturedGroup
          label={globalTools.label}
          testId="pdf-featured-global-tools-group"
        >
          <PdfGlobalToolsOverflow {...globalTools} />
        </FeaturedGroup>
      ) : null}
      {status ? (
        <span
          className="pdf-viewer__snapshot-toast office-pop-in"
          role="status"
          data-testid={statusTestId ?? "pdf-featured-action-status"}
        >
          {status}
        </span>
      ) : null}
    </>
  );
}
