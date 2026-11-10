"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useRef } from "react";

import PdfRibbonButton from "./PdfRibbonButton";

import { PdfPageInsertBlankIcon } from "./PdfPageRibbonIcons";
import { PdfSurfaceRibbonGroup } from "./PdfSurfaceRibbon";
import styles from "./PdfPageRibbonTab.module.css";

export const PDF_PAGE_INSERT_BLANK_CONTROL_ID =
  "PDFCTL.RIBBON.PAGE.INSERT_BLANK";

export interface PdfPageInsertBlankRibbonControlProps {
  disabled?: boolean;
  disabledReason?: string;
  insertBlankLabel: string;
  insertBlankLoadingLabel?: string;
  insertBlankOpen?: boolean;
  loading?: boolean;
  onInsertBlank: (trigger: HTMLButtonElement) => void;
}

export interface PdfPageRibbonTabProps extends PdfPageInsertBlankRibbonControlProps {
  pageStructureLabel: string;
}

/**
 * Page structure control 01. Exported separately so an integration owner can
 * replace the legacy inline host button with one exact hunk while the full
 * PAGE group remains owned here.
 */
export function PdfPageInsertBlankRibbonControl({
  disabled = false,
  disabledReason,
  insertBlankLabel,
  insertBlankLoadingLabel,
  insertBlankOpen = false,
  loading = false,
  onInsertBlank,
}: PdfPageInsertBlankRibbonControlProps) {
  const commandDisabled = disabled || loading;
  const state = loading ? "loading" : insertBlankOpen ? "open" : "idle";
  const triggerRef = useRef<HTMLButtonElement>(null);

  return (
    <>
      <PdfRibbonButton
        aria-busy={loading || undefined}
        aria-expanded={insertBlankOpen}
        aria-haspopup="dialog"
        buttonRef={triggerRef}
        className={styles.insertBlankCommand}
        data-control-id={PDF_PAGE_INSERT_BLANK_CONTROL_ID}
        data-disabled-reason={
          commandDisabled
            ? loading
              ? "pdf.page.insert-blank.loading"
              : disabledReason
            : undefined
        }
        data-loading={loading ? "true" : "false"}
        data-state={state}
        disabled={commandDisabled}
        label={insertBlankLabel}
        onClick={() => {
          if (triggerRef.current) onInsertBlank(triggerRef.current);
        }}
        testId="pdf-page-insert-blank"
      >
        <span className={styles.iconSlot} aria-hidden="true">
          <PdfPageInsertBlankIcon
            size={24}
            state={
              loading ? "loading" : commandDisabled ? "disabled" : "default"
            }
          />
          {loading ? <span className={styles.loadingSpinner} /> : null}
        </span>
      </PdfRibbonButton>
      {loading ? (
        <span
          className={styles.visuallyHidden}
          role="status"
          aria-live="polite"
        >
          {insertBlankLoadingLabel ?? insertBlankLabel}
        </span>
      ) : null}
    </>
  );
}

/**
 * WPS-compatible PDF PAGE ribbon, populated strictly left-to-right from the
 * frozen Control atlas. This first production slice owns only Page structure
 * control 01 (Insert Blank); subsequent controls join this module in atlas
 * order instead of being improvised in either PDF host.
 */
export default function PdfPageRibbonTab({
  pageStructureLabel,
  ...insertBlankProps
}: PdfPageRibbonTabProps) {
  return (
    <PdfSurfaceRibbonGroup
      label={pageStructureLabel}
      tab="insert page"
      testId="pdf-page-ribbon-page-structure"
    >
      <PdfPageInsertBlankRibbonControl {...insertBlankProps} />
    </PdfSurfaceRibbonGroup>
  );
}
