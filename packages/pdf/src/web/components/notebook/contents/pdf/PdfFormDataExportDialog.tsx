"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { useState, type RefObject } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import type { PdfFormDataFormat } from "@/lib/pdf/form-export";

interface PdfFormDataExportDialogProps {
  readonly open: boolean;
  readonly value: PdfFormDataFormat;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
  readonly onApply: (format: PdfFormDataFormat) => void;
  readonly onCancel: () => void;
}

const FORMATS: readonly PdfFormDataFormat[] = ["csv", "xfdf"];

function PdfFormDataExportDialogSession({
  value,
  returnFocusRef,
  onApply,
  onCancel,
}: Omit<PdfFormDataExportDialogProps, "open">): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const [draft, setDraft] = useState<PdfFormDataFormat>(value);

  return (
    <OfficeDialogShell
      open
      title={t("pdf.formDataExport.title")}
      description={t("pdf.formDataExport.description")}
      defaultActionLabel={t("pdf.formDataExport.export")}
      cancelActionLabel={t("pdf.formDataExport.cancel")}
      initialFocus="cancel"
      returnFocusRef={returnFocusRef}
      onDefaultAction={() => onApply(draft)}
      onCancel={onCancel}
      testIdPrefix="pdf-form-data-export"
    >
      <fieldset className="grid gap-2">
        <legend className="sr-only">{t("pdf.formDataExport.format")}</legend>
        {FORMATS.map((format) => (
          <label
            key={format}
            className={`${parts.check} cursor-pointer min-h-11`}
          >
            <span className="relative mt-0.5 size-4 shrink-0">
              <input
                type="radio"
                name="pdf-form-data-export-format"
                value={format}
                checked={draft === format}
                className={`${parts.markInput} absolute -left-3 -top-3 size-11 cursor-pointer`}
                data-testid={`pdf-form-data-export-${format}`}
                onChange={() => setDraft(format)}
              />
              <span
                aria-hidden
                className={`${parts.mark} pointer-events-none absolute inset-0 inline-flex size-4 items-center justify-center`} data-shape="radio"
              />
            </span>
            <span className="grid min-w-0 gap-0.5">
              <span className="font-medium">{format.toUpperCase()}</span>
              <span className={parts.hint}>
                {format === "csv"
                  ? t("pdf.formDataExport.csvHint")
                  : t("pdf.formDataExport.xfdfHint")}
              </span>
            </span>
          </label>
        ))}
      </fieldset>
    </OfficeDialogShell>
  );
}

export default function PdfFormDataExportDialog(
  props: PdfFormDataExportDialogProps,
): React.JSX.Element | null {
  if (!props.open) return null;
  return (
    <PdfFormDataExportDialogSession
      key={props.value}
      value={props.value}
      returnFocusRef={props.returnFocusRef}
      onApply={props.onApply}
      onCancel={props.onCancel}
    />
  );
}
