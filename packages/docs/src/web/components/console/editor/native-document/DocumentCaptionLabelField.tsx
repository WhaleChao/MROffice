"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Label part of Word's Caption dialog (References > Insert Caption):
 * the caption as it will read ("Figure 1 ..."), the Label list — Word's
 * built-in labels in the app language plus the custom ones — and New
 * Label... / Delete Label. A custom label numbers on its own (its own SEQ
 * sequence); a built-in label cannot be deleted; deleting a label leaves the
 * captions written with it untouched.
 */

import { useMemo, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  MAX_CAPTION_LABEL_NAME_LENGTH,
  captionInSequence,
  captionLabelNameProblem,
  captionLabelOptionKey,
  captionLabelOptions,
  captionPrefix,
  captionSequenceKey,
  sanitizeCaptionLabelName,
  type CaptionKind,
  type CaptionNumberFormat,
  type CaptionSeparator,
  type DocCaption,
} from "@mrnote/core/web/components/console/editor/native-document/captions.ts";
import {
  deleteCaptionLabel,
  readCaptionLabelStore,
  saveCaptionLabel,
} from "./captionLabelStore";

export interface CaptionLabelChoice {
  kind: CaptionKind;
  labelName: string | null;
}

export interface DocumentCaptionLabelFieldProps {
  /** The document's captions when the dialog opened. */
  readonly captions: readonly DocCaption[];
  readonly kind: CaptionKind;
  readonly labelName: string | null;
  readonly onChange: (choice: CaptionLabelChoice) => void;
  /** The caption being edited (its number is kept in the preview). */
  readonly editingCaptionId?: string | null;
  /**
   * How many of `captions` come before the place the new caption goes: its
   * number follows theirs, as the field Word inserts there reads.
   */
  readonly insertionIndex?: number;
  readonly text: string;
  readonly numberFormat: CaptionNumberFormat;
  readonly excludeLabel: boolean;
  readonly separator: CaptionSeparator;
}

export function DocumentCaptionLabelField({
  captions,
  kind,
  labelName,
  onChange,
  editingCaptionId = null,
  insertionIndex,
  text,
  numberFormat,
  excludeLabel,
  separator,
}: DocumentCaptionLabelFieldProps) {
  const t = useTranslations("console-notebooks");
  const [store, setStore] = useState(readCaptionLabelStore);
  const [naming, setNaming] = useState(false);
  const [nameDraft, setNameDraft] = useState("");
  const [nameError, setNameError] = useState<string | null>(null);

  const selectedKey = captionLabelOptionKey(kind, labelName);
  const options = useMemo(() => {
    const listed = captionLabelOptions(captions, store.saved, store.deleted);
    const selectedName = sanitizeCaptionLabelName(labelName);
    // The caption's own label stays choosable even after it was deleted.
    if (selectedName && !listed.some((option) => option.key === selectedKey)) {
      listed.push({
        key: selectedKey,
        kind,
        labelName: selectedName,
        custom: kind === "custom",
      });
    }
    return listed;
  }, [captions, kind, labelName, selectedKey, store]);
  const selected = options.find((option) => option.key === selectedKey) ?? null;

  const previewNumber = useMemo(() => {
    const sequence = captionSequenceKey(kind, labelName);
    const editing = editingCaptionId
      ? captions.find((caption) => caption.id === editingCaptionId)
      : undefined;
    if (editing && captionInSequence(editing, sequence)) return editing.number;
    const before =
      typeof insertionIndex === "number" ? captions.slice(0, insertionIndex) : captions;
    // The number after the last one of the label before this place: a
    // chapter-numbered sequence restarts (Figure 2-1), so it is not a count.
    const previous = before.filter((caption) => captionInSequence(caption, sequence)).at(-1);
    return (previous?.number ?? 0) + 1;
  }, [captions, editingCaptionId, insertionIndex, kind, labelName]);
  const prefix = captionPrefix(kind, previewNumber, {
    labelName,
    numberFormat,
    excludeLabel,
  });
  const preview = text.trim() ? `${prefix}${separator}${text.trim()}` : prefix;

  const confirmNewLabel = () => {
    const problem = captionLabelNameProblem(nameDraft, options);
    if (problem === "exists") {
      // Word selects the label that is already in the list.
      const existing = options.find(
        (option) =>
          option.labelName.toLowerCase() ===
          sanitizeCaptionLabelName(nameDraft).toLowerCase(),
      );
      if (existing) {
        onChange({ kind: existing.kind, labelName: existing.labelName });
        setNaming(false);
        setNameDraft("");
        setNameError(null);
        return;
      }
    }
    if (problem) {
      setNameError(
        t(
          problem === "tooLong"
            ? "toolbar.captionNewLabelTooLong"
            : problem === "exists"
              ? "toolbar.captionNewLabelExists"
              : "toolbar.captionNewLabelEmpty",
        ),
      );
      return;
    }
    const name = sanitizeCaptionLabelName(nameDraft);
    setStore((current) => saveCaptionLabel(current, name));
    onChange({ kind: "custom", labelName: name });
    setNaming(false);
    setNameDraft("");
    setNameError(null);
  };

  return (
    <>
      <output
        className={`${officeDialogParts.well} document-caption-preview`}
        aria-label={t("toolbar.captionPreview")}
        data-testid="document-caption-preview"
      >
        {preview}
      </output>
      <div className={officeDialogParts.row}>
        <label className={officeDialogParts.field}>
          {t("toolbar.captionKindLabel")}
          <select
            data-testid="document-caption-kind"
            aria-label={t("toolbar.captionKindLabel")}
            value={selectedKey}
            onChange={(event) => {
              const next = options.find(
                (option) => option.key === event.target.value,
              );
              if (next) onChange({ kind: next.kind, labelName: next.labelName });
            }}
          >
            {options.map((option) => (
              <option key={option.key} value={option.key}>
                {option.labelName}
              </option>
            ))}
          </select>
        </label>
        <button
          type="button"
          className={officeDialogParts.button}
          data-testid="document-caption-new-label"
          aria-expanded={naming}
          onClick={() => {
            setNaming(true);
            setNameError(null);
          }}
        >
          {t("toolbar.captionNewLabel")}
        </button>
        <button
          type="button"
          className={officeDialogParts.button}
          data-testid="document-caption-delete-label"
          disabled={!selected?.custom}
          onClick={() => {
            if (!selected?.custom) return;
            setStore((current) =>
              deleteCaptionLabel(current, selected.labelName),
            );
            const fallback = options[0];
            onChange({ kind: fallback.kind, labelName: fallback.labelName });
          }}
        >
          {t("toolbar.captionDeleteLabel")}
        </button>
      </div>
      {naming ? (
        <div
          className={officeDialogParts.row}
          data-testid="document-caption-new-label-form"
        >
          <label className={officeDialogParts.field}>
            {t("toolbar.captionNewLabelName")}
            <input
              type="text"
              autoFocus
              maxLength={MAX_CAPTION_LABEL_NAME_LENGTH}
              data-testid="document-caption-new-label-name"
              aria-label={t("toolbar.captionNewLabelName")}
              aria-invalid={nameError ? true : undefined}
              value={nameDraft}
              onChange={(event) => {
                setNameDraft(event.target.value);
                setNameError(null);
              }}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  // Enter adds the label; it must not also press the
                  // dialog's OK.
                  event.preventDefault();
                  event.stopPropagation();
                  confirmNewLabel();
                } else if (event.key === "Escape") {
                  event.preventDefault();
                  event.stopPropagation();
                  setNaming(false);
                  setNameError(null);
                }
              }}
            />
          </label>
          <button
            type="button"
            className={officeDialogParts.button}
            data-testid="document-caption-new-label-confirm"
            onClick={confirmNewLabel}
          >
            {t("toolbar.dialogOk")}
          </button>
          <button
            type="button"
            className={officeDialogParts.button}
            data-testid="document-caption-new-label-cancel"
            onClick={() => {
              setNaming(false);
              setNameDraft("");
              setNameError(null);
            }}
          >
            {t("toolbar.pageSetupCancel")}
          </button>
        </div>
      ) : null}
      {nameError ? (
        <p
          className={officeDialogParts.error}
          role="alert"
          data-testid="document-caption-new-label-error"
        >
          {nameError}
        </p>
      ) : null}
    </>
  );
}
