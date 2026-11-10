"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { RefObject } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";

export interface PdfDangerousActionEvidence {
  readonly kind: string;
  readonly location: string;
}

interface PdfDangerousActionSafeCopyDialogProps {
  readonly open: boolean;
  readonly actions: readonly PdfDangerousActionEvidence[];
  readonly busy: boolean;
  readonly error: string | null;
  readonly onCancel: () => void;
  readonly onConfirm: () => void;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
}

export default function PdfDangerousActionSafeCopyDialog({
  open,
  actions,
  busy,
  error,
  onCancel,
  onConfirm,
  returnFocusRef,
}: PdfDangerousActionSafeCopyDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  return (
    <OfficeDialogShell
      open={open}
      title={t("pdf.safeCopy.title")}
      description={t("pdf.safeCopy.description")}
      defaultActionLabel={t("pdf.safeCopy.apply")}
      cancelActionLabel={t("pdf.safeCopy.cancel")}
      busyActionLabel={t("pdf.safeCopy.verifying")}
      defaultActionDisabled={busy || actions.length === 0}
      busy={busy}
      initialFocus="cancel"
      onDefaultAction={onConfirm}
      onCancel={onCancel}
      returnFocusRef={returnFocusRef}
      testIdPrefix="pdf-dangerous-action-safe-copy"
      className="sm:max-w-[640px]"
    >
      <div className="grid gap-3">
        <p className={parts.message} data-tone="warning">
          {t("pdf.safeCopy.existingActions", { count: actions.length })}
        </p>
        <ul
          className={`${parts.list} max-h-64 space-y-2 overflow-auto`}
          data-testid="pdf-dangerous-action-safe-copy-locations"
        >
          {actions.map((action, index) => (
            <li
              key={`${action.location}:${action.kind}:${index}`}
              className={`${parts.well} grid gap-0.5`}
            >
              <strong>{action.kind}</strong>
              <code className={`${parts.hint} break-all`}>
                {action.location}
              </code>
            </li>
          ))}
        </ul>
        <p className={parts.hint}>{t("pdf.safeCopy.rescanHint")}</p>
        {error ? (
          <p
            className={parts.message}
            role="alert"
            data-testid="pdf-dangerous-action-safe-copy-error"
          >
            {error}
          </p>
        ) : null}
      </div>
    </OfficeDialogShell>
  );
}
