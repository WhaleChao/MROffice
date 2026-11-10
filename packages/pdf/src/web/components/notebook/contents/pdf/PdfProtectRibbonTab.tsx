"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  BadgeCheck,
  FileCheck2,
  LockKeyhole,
  ScanLine,
  ShieldAlert,
  ShieldCheck,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useRef, type ReactNode } from "react";

import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeRibbonGroup,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";

import styles from "./PdfProtectRibbonTab.module.css";

export interface PdfProtectRibbonCommand {
  active?: boolean;
  busy?: boolean;
  disabled?: boolean;
  disabledReason?: string;
  label: string;
  onInvoke?: (trigger: HTMLButtonElement | null) => void;
  testId: string;
}

export interface PdfProtectRibbonLabels {
  encryption: string;
  redaction: string;
  validation: string;
  safety: string;
  encryptCopy: string;
  markRedactions: string;
  applyRedactions: string;
  safeCopy: string;
  archivalCopy: string;
  signatureStatus: string;
  actionStatus: string;
}

export interface PdfProtectRibbonStatus {
  signature: string;
  actions: string;
  actionRisk: boolean;
}

function commandLabel(command: PdfProtectRibbonCommand): string {
  return command.disabled && command.disabledReason
    ? `${command.label}: ${command.disabledReason}`
    : command.label;
}

function ProtectCommand({
  command,
  icon,
}: {
  command: PdfProtectRibbonCommand;
  icon: ReactNode;
}) {
  const triggerRef = useRef<HTMLButtonElement>(null);
  return (
    <OfficeHomeCommand
      activation="click"
      active={command.active}
      buttonRef={triggerRef}
      disabled={command.disabled || command.busy || !command.onInvoke}
      displayLabel={command.label}
      icon={icon}
      label={commandLabel(command)}
      onClick={() => command.onInvoke?.(triggerRef.current)}
      testId={command.testId}
      variant="large"
    />
  );
}

export default function PdfProtectRibbonTab({
  visible,
  ribbon = "protect",
  labels,
  commands,
  status,
}: {
  visible: boolean;
  ribbon?: string;
  labels: PdfProtectRibbonLabels;
  commands: {
    encrypt: PdfProtectRibbonCommand;
    mark: PdfProtectRibbonCommand;
    apply: PdfProtectRibbonCommand;
    safeCopy: PdfProtectRibbonCommand;
    archival: PdfProtectRibbonCommand;
  };
  status: PdfProtectRibbonStatus;
}) {
  if (!visible) return null;
  return (
    <>
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.encryption}
        size="content"
        testId="pdf-protect-encryption-group"
      >
        <ProtectCommand
          command={commands.encrypt}
          icon={<LockKeyhole size={24} aria-hidden />}
        />
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.redaction}
        size="content"
        testId="pdf-protect-redaction-group"
      >
        <OfficeHomeCommandRow>
          <ProtectCommand
            command={commands.mark}
            icon={<ScanLine size={24} aria-hidden />}
          />
          <ProtectCommand
            command={commands.apply}
            icon={<ShieldCheck size={24} aria-hidden />}
          />
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.safety}
        size="content"
        testId="pdf-protect-safety-group"
      >
        <ProtectCommand
          command={commands.safeCopy}
          icon={<ShieldAlert size={24} aria-hidden />}
        />
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.validation}
        size="content"
        testId="pdf-protect-validation-group"
      >
        <ProtectCommand
          command={commands.archival}
          icon={<FileCheck2 size={24} aria-hidden />}
        />
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.signatureStatus}
        size="content"
        className={styles.statusGroup}
        testId="pdf-protect-status-group"
      >
        <div className={styles.statusGrid}>
          <div
            className={styles.statusItem}
            data-testid="pdf-protect-signature-status"
          >
            <BadgeCheck size={16} aria-hidden />
            <span>
              <strong>{labels.signatureStatus}</strong>
              <small>{status.signature}</small>
            </span>
          </div>
          <div
            className={`${styles.statusItem}${status.actionRisk ? ` ${styles.risk}` : ""}`}
            data-testid="pdf-protect-action-status"
          >
            {status.actionRisk ? (
              <ShieldAlert size={16} aria-hidden />
            ) : (
              <ShieldCheck size={16} aria-hidden />
            )}
            <span>
              <strong>{labels.actionStatus}</strong>
              <small>{status.actions}</small>
            </span>
          </div>
        </div>
      </OfficeHomeRibbonGroup>
    </>
  );
}
