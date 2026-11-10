"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  ExternalLink,
  FileText,
  MoveDiagonal2,
  Pencil,
  ShieldAlert,
  Trash2,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";

import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeRibbonGroup,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";

import styles from "./PdfLinkContextRibbon.module.css";
import type { PdfRibbonTranslator } from "./pdfRibbonTranslator";

export type PdfLinkContextKind = "url" | "page";
export type PdfLinkContextSource = "managed" | "native";

export interface PdfLinkContextRibbonModel {
  kind: PdfLinkContextKind;
  source: PdfLinkContextSource;
  page: number;
  destination: string;
  width: number;
  height: number;
  busy: boolean;
  openEnabled: boolean;
  editEnabled: boolean;
  editDisabledReason?: string;
  deleteEnabled: boolean;
  deleteDisabledReason?: string;
}

export interface PdfLinkContextRibbonLabels {
  link: string;
  open: string;
  edit: string;
  destination: string;
  positionSize: string;
  actions: string;
  webLink: string;
  pageLink: string;
  managedLink: string;
  nativeLink: string;
  openLink: string;
  editLink: string;
  removeLink: string;
  backToEditing: string;
  pageNumber: (page: number) => string;
}

export function createPdfLinkContextRibbonLabels(
  t: PdfRibbonTranslator,
): PdfLinkContextRibbonLabels {
  return {
    link: t("pdfContextRibbon.link.link"),
    open: t("pdfContextRibbon.link.open"),
    edit: t("pdfContextRibbon.link.edit"),
    destination: t("pdfContextRibbon.link.destination"),
    positionSize: t("pdfContextRibbon.link.positionSize"),
    actions: t("pdfContextRibbon.link.actions"),
    webLink: t("pdfContextRibbon.link.webLink"),
    pageLink: t("pdfContextRibbon.link.pageLink"),
    managedLink: t("pdfContextRibbon.link.managedLink"),
    nativeLink: t("pdfContextRibbon.link.nativeLink"),
    openLink: t("pdfContextRibbon.link.openLink"),
    editLink: t("pdfContextRibbon.link.editLink"),
    removeLink: t("pdfContextRibbon.link.removeLink"),
    backToEditing: t("pdfContextRibbon.link.backToEditing"),
    pageNumber: (page) => t("pdfContextRibbon.link.pageNumber", { page }),
  };
}

export function PdfLinkContextRibbon({
  labels,
  model,
  onOpen,
  onEdit,
  onDelete,
  onExit,
}: {
  labels: PdfLinkContextRibbonLabels;
  model: PdfLinkContextRibbonModel;
  onOpen: () => void;
  onEdit: () => void;
  onDelete: () => void;
  onExit: () => void;
}) {
  const openDisabled = model.busy || !model.openEnabled;
  const editDisabled = model.busy || !model.editEnabled;
  const deleteDisabled = model.busy || !model.deleteEnabled;
  return (
    <>
      <OfficeHomeRibbonGroup
        label={labels.link}
        ribbon="link-format"
        controlId="PDFCTL.CONTEXT.LINK.IDENTITY"
        className={styles.group}
        testId="pdf-link-context-link-group"
      >
        {/* No identity panel (WPS: none); the destination row says what
            the link is. */}
        <div
          className={styles.destination}
          data-testid="pdf-link-context-destination"
          title={model.destination}
        >
          {model.kind === "url" ? (
            <ExternalLink aria-hidden size={16} />
          ) : (
            <FileText aria-hidden size={16} />
          )}
          <span>{model.destination}</span>
        </div>
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.LINK.OPEN.COMMAND"
          disabled={openDisabled}
          icon={
            model.kind === "url" ? (
              <ExternalLink size={18} />
            ) : (
              <FileText size={18} />
            )
          }
          label={labels.openLink}
          onClick={onOpen}
          testId="pdf-link-context-open"
        />
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.LINK.EDIT.COMMAND"
          disabled={editDisabled}
          icon={editDisabled ? <ShieldAlert size={18} /> : <Pencil size={18} />}
          label={model.editDisabledReason ?? labels.editLink}
          displayLabel={labels.editLink}
          onClick={onEdit}
          testId="pdf-link-context-edit"
        />
      </OfficeHomeRibbonGroup>

      {/* 操作: the link's place on its page, delete, and 退出编辑 (every
          contextual tab ends with it). Two groups where six held one control
          each (audit A7). */}
      <OfficeHomeRibbonGroup
        label={labels.actions}
        ribbon="link-format"
        controlId="PDFCTL.CONTEXT.LINK.ACTIONS"
        className={styles.group}
        testId="pdf-link-context-actions-group"
      >
        <OfficeHomeCommandRow testId="pdf-link-context-size-readout">
          <span className={styles.measurement}>
            {model.width.toFixed(1)} × {model.height.toFixed(1)}%
          </span>
          <span className={styles.page}>{labels.pageNumber(model.page)}</span>
          <MoveDiagonal2 aria-hidden size={15} />
        </OfficeHomeCommandRow>
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.LINK.REMOVE"
          disabled={deleteDisabled}
          icon={
            deleteDisabled ? <ShieldAlert size={18} /> : <Trash2 size={18} />
          }
          label={model.deleteDisabledReason ?? labels.removeLink}
          displayLabel={labels.removeLink}
          onClick={onDelete}
          testId="pdf-link-context-delete"
        />
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.LINK.EXIT"
          disabled={model.busy}
          displayLabel={labels.backToEditing}
          icon={<OfficeSemanticIcon semanticId="edit.exit" size={24} />}
          label={labels.backToEditing}
          onClick={onExit}
          testId="pdf-link-context-exit"
          variant="large"
        />
      </OfficeHomeRibbonGroup>
    </>
  );
}
