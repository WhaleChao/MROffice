"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Check } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { useEffect, useRef, useState } from "react";
import { useOpenedOnce } from "@mrnote/core/web/lib/use-opened-once.ts";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { isNativePdfFormExportValue } from "../../../../lib/pdf/form-field-authoring";

export interface PdfRadioGroupOptionDraft {
  readonly key: string;
  readonly originalValue: string;
}

export interface PdfRadioGroupOptionsApplyValue {
  readonly replacementValues: readonly string[];
  readonly selectedIndex: number | null;
  readonly offToggleable: boolean;
}

type PdfRadioGroupOptionsDialogProps = {
  open: boolean;
  fieldName: string;
  options: readonly PdfRadioGroupOptionDraft[];
  selectedIndex: number | null;
  offToggleable: boolean;
  busy: boolean;
  /** Why the last save did not reach the PDF, shown inside this dialog. */
  submitError?: string | null;
  onApply: (value: PdfRadioGroupOptionsApplyValue) => Promise<boolean>;
  onCancel: () => void;
};

export function PdfRadioGroupOptionsDialog(props: PdfRadioGroupOptionsDialogProps) {
  if (!useOpenedOnce(props.open)) return null;
  return <PdfRadioGroupOptionsDialogBody {...props} />;
}

function PdfRadioGroupOptionsDialogBody({
  open,
  fieldName,
  options,
  selectedIndex: initialSelectedIndex,
  offToggleable: initialOffToggleable,
  busy,
  submitError = null,
  onApply,
  onCancel,
}: PdfRadioGroupOptionsDialogProps) {
  const t = useTranslations("console-notebooks");
  const firstInputRef = useRef<HTMLInputElement>(null);
  const submitInFlightRef = useRef(false);
  const [replacementValues, setReplacementValues] = useState<string[]>([]);
  const [selectedIndex, setSelectedIndex] = useState<number | null>(null);
  const [offToggleable, setOffToggleable] = useState(false);

  useEffect(() => {
    if (!open) return;
    setReplacementValues(options.map((option) => option.originalValue));
    setSelectedIndex(initialSelectedIndex);
    setOffToggleable(initialOffToggleable);
  }, [initialOffToggleable, initialSelectedIndex, open, options]);

  const validValues =
    replacementValues.length === options.length &&
    replacementValues.length >= 2 &&
    replacementValues.every(isNativePdfFormExportValue) &&
    new Set(replacementValues).size === replacementValues.length;
  const validSelection =
    selectedIndex === null
      ? offToggleable
      : Number.isSafeInteger(selectedIndex) &&
        selectedIndex >= 0 &&
        selectedIndex < replacementValues.length;
  const canApply = validValues && validSelection && !busy;

  const apply = async () => {
    if (!canApply || submitInFlightRef.current) return;
    submitInFlightRef.current = true;
    try {
      await onApply({
        replacementValues,
        selectedIndex,
        offToggleable,
      });
    } finally {
      submitInFlightRef.current = false;
    }
  };

  return (
    <OfficeDialogShell
      open={open}
      title={t("pdf.radioGroup.title")}
      description={t("pdf.radioGroup.description", { fieldName })}
      defaultActionLabel={t("pdf.radioGroup.apply")}
      cancelActionLabel={t("pdf.radioGroup.cancel")}
      onDefaultAction={() => void apply()}
      onCancel={onCancel}
      defaultActionDisabled={!canApply}
      busy={busy}
      busyActionLabel={t("pdf.radioGroup.saving")}
      initialFocusRef={firstInputRef}
      testIdPrefix="pdf-radio-group-options"
    >
      <div className="grid gap-3" data-testid="pdf-radio-options-list">
        {options.map((option, index) => (
          <label
            key={option.key}
            className={`${parts.field} grid-cols-[44px_minmax(0,1fr)] min-h-11`}
          >
            <span className="relative size-11 shrink-0">
              <input
                type="radio"
                name="pdf-radio-group-selected-option"
                checked={selectedIndex === index}
                onChange={() => setSelectedIndex(index)}
                disabled={busy}
                aria-label={t("pdf.radioGroup.selectOption", {
                  index: index + 1,
                })}
                className={`${parts.markInput} absolute inset-0 size-11 cursor-pointer`}
                data-testid="pdf-radio-option-selected"
              />
              <span
                aria-hidden
                className={`${parts.mark} pointer-events-none absolute left-3.5 top-3.5 inline-flex size-4 items-center justify-center`} data-shape="radio"
              />
            </span>
            <input
              ref={index === 0 ? firstInputRef : undefined}
              value={replacementValues[index] ?? ""}
              maxLength={127}
              onChange={(event) =>
                setReplacementValues((current) =>
                  current.map((value, valueIndex) =>
                    valueIndex === index ? event.target.value : value,
                  ),
                )
              }
              disabled={busy}
              autoComplete="off"
              aria-label={t("pdf.radioGroup.optionLabel", { index: index + 1 })}
              className="min-w-0"
              data-testid="pdf-radio-option-value"
            />
          </label>
        ))}
      </div>
      <label className={`${parts.check} mt-3 min-h-11`}>
        <span className="relative size-11 shrink-0">
          <input
            type="checkbox"
            checked={offToggleable}
            onChange={(event) => {
              const next = event.target.checked;
              setOffToggleable(next);
              if (!next && selectedIndex === null && options.length > 0) {
                setSelectedIndex(0);
              }
            }}
            disabled={busy}
            className={`${parts.markInput} absolute inset-0 size-11 cursor-pointer`}
            data-testid="pdf-radio-option-off-toggleable"
          />
          <span
            aria-hidden
            className={`${parts.mark} pointer-events-none absolute left-3.5 top-3.5 inline-flex size-4 items-center justify-center`}
          >
            {offToggleable ? <Check size={12} strokeWidth={2.4} /> : null}
          </span>
        </span>
        {t("pdf.radioGroup.allowClear")}
      </label>
      {offToggleable ? (
        <button
          type="button"
          onClick={() => setSelectedIndex(null)}
          disabled={busy}
          className={`${parts.compactButton} mt-2`}
          data-testid="pdf-radio-option-clear-selection"
        >
          {t("pdf.radioGroup.noOption")}
        </button>
      ) : null}
      {!validValues ? (
        <p
          role="alert"
          className={`${parts.error} mt-3`}
          data-testid="pdf-radio-options-error"
        >
          {t("pdf.radioGroup.invalidOptions")}
        </p>
      ) : submitError ? (
        <p
          role="alert"
          className={`${parts.error} mt-3`}
          data-testid="pdf-radio-options-submit-error"
        >
          {submitError}
        </p>
      ) : null}
    </OfficeDialogShell>
  );
}
