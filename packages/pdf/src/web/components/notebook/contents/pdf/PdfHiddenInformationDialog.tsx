"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import {
  PDF_HIDDEN_INFORMATION_KINDS,
  type PdfHiddenInformationKind,
  type PdfHiddenInformationReport,
} from "../../../../lib/pdf/hidden-information";

export interface PdfHiddenInformationDialogProps {
  open: boolean;
  /** Reads what the file holds (once, as the dialog opens). */
  inspect: () => Promise<PdfHiddenInformationReport>;
  busy?: boolean;
  error?: string | null;
  onConfirm: (kinds: PdfHiddenInformationKind[]) => void;
  onCancel: () => void;
}

/**
 * 删除隐藏信息 (ONLYOFFICE, Acrobat): what the file carries that its pages
 * do not show, listed by kind with how much of each there is; the kinds
 * ticked are removed and the file is written again without them.
 */
export default function PdfHiddenInformationDialog({
  open,
  inspect,
  busy = false,
  error = null,
  onConfirm,
  onCancel,
}: PdfHiddenInformationDialogProps) {
  const commonT = useTranslations("common");
  const t = useTranslations("console-notebooks");
  const [report, setReport] = useState<PdfHiddenInformationReport | null>(null);
  const [inspectFailed, setInspectFailed] = useState(false);
  const [chosen, setChosen] = useState<ReadonlySet<PdfHiddenInformationKind>>(new Set());
  const inspectRef = useRef(inspect);
  inspectRef.current = inspect;
  useEffect(() => {
    if (!open) return;
    let live = true;
    setReport(null);
    setInspectFailed(false);
    inspectRef.current().then(
      (next) => {
        if (!live) return;
        setReport(next);
        setChosen(new Set(PDF_HIDDEN_INFORMATION_KINDS.filter((kind) => next[kind] > 0)));
      },
      () => {
        if (live) setInspectFailed(true);
      },
    );
    return () => {
      live = false;
    };
  }, [open]);
  const toggle = (kind: PdfHiddenInformationKind) =>
    setChosen((current) => {
      const next = new Set(current);
      if (next.has(kind)) next.delete(kind);
      else next.add(kind);
      return next;
    });
  const found = report ? PDF_HIDDEN_INFORMATION_KINDS.filter((kind) => report[kind] > 0) : [];
  return (
    <OfficeDialogShell
      open={open}
      title={t("pdfHidden.title")}
      description={t("pdfHidden.description")}
      defaultActionLabel={t("pdfHidden.remove")}
      cancelActionLabel={commonT("button.cancel")}
      onDefaultAction={() => {
        if (chosen.size > 0 && !busy) onConfirm([...chosen]);
      }}
      onCancel={onCancel}
      defaultActionDisabled={busy || chosen.size === 0}
      className="sm:max-w-[460px]"
      testIdPrefix="pdf-hidden-information"
    >
      <div style={{ display: "grid", gap: 10 }} data-testid="pdf-hidden-information-list">
        {inspectFailed ? (
          <p role="alert">{t("pdfHidden.inspectFailed")}</p>
        ) : !report ? (
          <p role="status">{t("pdfHidden.inspecting")}</p>
        ) : found.length === 0 ? (
          <p role="status" data-testid="pdf-hidden-information-none">{t("pdfHidden.none")}</p>
        ) : (
          found.map((kind) => (
            <label key={kind} style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <input
                type="checkbox"
                checked={chosen.has(kind)}
                onChange={() => toggle(kind)}
                data-testid={`pdf-hidden-information-${kind}`}
              />
              <span>{t(`pdfHidden.kind.${kind}`)}</span>
              <span style={{ marginLeft: "auto", opacity: 0.7 }}>{report[kind]}</span>
            </label>
          ))
        )}
        {error ? <p role="alert">{error}</p> : null}
      </div>
    </OfficeDialogShell>
  );
}
