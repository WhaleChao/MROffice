"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useMemo, useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { rotationTargetPages, type PdfPageRotationScope } from "../../../../lib/pdf/page-rotation";
import { formatPdfPageScopePages, PdfPageScopeChoice } from "./PdfPageScopeChoice";

export interface PdfRedactPagesDialogProps {
  open: boolean;
  currentPage: number;
  totalPages: number;
  busy?: boolean;
  /** The pages to mark, each whole page one redaction mark. */
  onConfirm: (pages: number[]) => void;
  onCancel: () => void;
}

/**
 * 页面涂黑 (ONLYOFFICE's Redact › Page, Acrobat's Redact Pages): whole pages
 * marked for redaction; 应用涂黑 then removes everything on them, as it does
 * under any mark.
 */
export default function PdfRedactPagesDialog({
  open,
  currentPage,
  totalPages,
  busy = false,
  onConfirm,
  onCancel,
}: PdfRedactPagesDialogProps) {
  const commonT = useTranslations("common");
  const t = useTranslations("console-notebooks");
  const [scope, setScope] = useState<PdfPageRotationScope>("page");
  const [rangeSpec, setRangeSpec] = useState(String(currentPage));
  const confirmInFlightRef = useRef(false);
  useEffect(() => {
    if (open) confirmInFlightRef.current = false;
  }, [open]);
  const pages = useMemo(
    () => rotationTargetPages(scope, currentPage, Math.max(1, totalPages), rangeSpec),
    [currentPage, rangeSpec, scope, totalPages],
  );
  const rangeInvalid = scope === "range" && pages.length === 0;
  const confirm = () => {
    if (pages.length === 0 || busy || confirmInFlightRef.current) return;
    confirmInFlightRef.current = true;
    onConfirm(pages);
  };
  return (
    <OfficeDialogShell
      open={open}
      title={t("pdfRedact.pages")}
      description={t("pdfRedact.pagesDescription")}
      defaultActionLabel={t("pdfRedact.pagesConfirm", { count: pages.length })}
      cancelActionLabel={commonT("button.cancel")}
      onDefaultAction={confirm}
      onCancel={onCancel}
      defaultActionDisabled={pages.length === 0 || busy}
      className="sm:max-w-[480px]"
      testIdPrefix="pdf-redact-pages"
    >
      <div style={{ display: "grid", gap: 14 }}>
        <PdfPageScopeChoice
          scope={scope}
          onScopeChange={setScope}
          rangeSpec={rangeSpec}
          onRangeSpecChange={setRangeSpec}
          rangeInvalid={rangeInvalid}
          currentPage={currentPage}
          totalPages={totalPages}
          testIdPrefix="pdf-redact"
        />
        <div role="status" data-testid="pdf-redact-pages-summary">
          {formatPdfPageScopePages(pages, t)}
        </div>
      </div>
    </OfficeDialogShell>
  );
}
