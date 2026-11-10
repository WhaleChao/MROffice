"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useMemo, useRef, useState, type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import type { ProofingLanguageValue } from "@mrnote/core/web/components/console/editor/native-document/proofingLanguageExtension.ts";

export interface DocumentProofingLanguageDialogProps {
  readonly canEdit?: boolean;
  readonly open: boolean;
  readonly locale: string;
  readonly initialValue: ProofingLanguageValue | null;
  readonly mixed: boolean;
  




  readonly noProofMixed?: boolean;
  /** Product name in the explanatory copy; shared Office editors may override it. */
  readonly productName?: string;
  readonly returnFocusRef: RefObject<HTMLElement | null>;
  readonly onSubmit: (
    value: ProofingLanguageValue,
    options?: { preserveLanguages?: boolean; preserveNoProof?: boolean },
  ) => boolean;
  readonly onClose: () => void;
}

const LANGUAGE_OPTIONS = [
  { key: "auto", val: null, eastAsia: null, bidi: null },
  { key: "en-US", val: "en-US", eastAsia: null, bidi: null },
  { key: "en-GB", val: "en-GB", eastAsia: null, bidi: null },
  { key: "fr-FR", val: "fr-FR", eastAsia: null, bidi: null },
  { key: "de-DE", val: "de-DE", eastAsia: null, bidi: null },
  { key: "es-ES", val: "es-ES", eastAsia: null, bidi: null },
  { key: "zh-CN", val: null, eastAsia: "zh-CN", bidi: null },
  { key: "zh-TW", val: null, eastAsia: "zh-TW", bidi: null },
  { key: "ja-JP", val: null, eastAsia: "ja-JP", bidi: null },
  { key: "ko-KR", val: null, eastAsia: "ko-KR", bidi: null },
  { key: "ar-SA", val: null, eastAsia: null, bidi: "ar-SA" },
] as const;

const MIXED_LANGUAGE_KEY = "__mixed__";
/** The document's own setting when it is not one of the listed presets. */
const CURRENT_LANGUAGE_KEY = "__current__";

/** Whether a stored value sets any language slot (it is not "automatic"). */
function setsLanguage(value: ProofingLanguageValue | null): boolean {
  return Boolean(value && (value.val || value.eastAsia || value.bidi));
}

/** The codes of an unlisted setting, as the dialog names it: "en-US / zh-CN". */
function currentLanguageLabel(value: ProofingLanguageValue): string {
  return [value.val, value.eastAsia, value.bidi].filter(Boolean).join(" / ");
}

function optionKey(value: ProofingLanguageValue | null): string {
  if (!value) return "auto";
  const match = LANGUAGE_OPTIONS.find(
    (option) =>
      option.val === value.val &&
      option.eastAsia === value.eastAsia &&
      option.bidi === value.bidi,
  );
  return match?.key ?? "auto";
}

export function DocumentProofingLanguageDialog({
  canEdit = true,
  open,
  initialValue,
  mixed,
  noProofMixed = false,
  productName = "MRNote",
  returnFocusRef,
  onSubmit,
  onClose,
}: DocumentProofingLanguageDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const copy = useMemo(
    () => ({
      title: t("toolbar.proofingLanguageDialog.title"),
      description: t("toolbar.proofingLanguageDialog.description", {
        productName,
      }),
      language: t("toolbar.proofingLanguageDialog.language"),
      mixed: t("toolbar.proofingLanguageDialog.mixed"),
      noProof: t("toolbar.proofingLanguageDialog.noProof"),
      apply: t("toolbar.proofingLanguageDialog.apply"),
      cancel: t("toolbar.proofingLanguageDialog.cancel"),
      stale: t("toolbar.proofingLanguageDialog.stale"),
      labels: {
        auto: t("toolbar.proofingLanguageDialog.auto"),
        "en-US": t("toolbar.proofingLanguageDialog.englishUnitedStates"),
        "en-GB": t("toolbar.proofingLanguageDialog.englishUnitedKingdom"),
        "fr-FR": t("toolbar.proofingLanguageDialog.frenchFrance"),
        "de-DE": t("toolbar.proofingLanguageDialog.germanGermany"),
        "es-ES": t("toolbar.proofingLanguageDialog.spanishSpain"),
        "zh-CN": t("toolbar.proofingLanguageDialog.chineseSimplified"),
        "zh-TW": t("toolbar.proofingLanguageDialog.chineseTraditional"),
        "ja-JP": t("toolbar.proofingLanguageDialog.japanese"),
        "ko-KR": t("toolbar.proofingLanguageDialog.korean"),
        "ar-SA": t("toolbar.proofingLanguageDialog.arabic"),
      } as Record<string, string>,
    }),
    [productName, t],
  );
  // A setting the list does not carry (a bilingual en-US + zh-CN run, it-IT)
  // is shown as itself and kept as itself: confirming without a change, or
  // only switching the checking off, used to rewrite it as "automatic" and
  // drop the languages the file had (OFL-053).
  const unlisted =
    !mixed && setsLanguage(initialValue) && optionKey(initialValue) === "auto";
  const [selectedKey, setSelectedKey] = useState(() =>
    mixed ? MIXED_LANGUAGE_KEY : unlisted ? CURRENT_LANGUAGE_KEY : optionKey(initialValue),
  );
  const [noProof, setNoProof] = useState(initialValue?.noProof === true);
  // Word's indeterminate box: until it is clicked, every run keeps its own.
  const [noProofTouched, setNoProofTouched] = useState(false);
  const noProofIndeterminate = noProofMixed && !noProofTouched;
  const [error, setError] = useState<string | null>(null);
  const selectRef = useRef<HTMLSelectElement | null>(null);
  // `indeterminate` is a DOM property only; set it whenever the box mounts
  // or the state changes (the shell may mount the body after this render).
  const noProofRef = useCallback(
    (box: HTMLInputElement | null) => {
      if (box) box.indeterminate = noProofIndeterminate;
    },
    [noProofIndeterminate],
  );

  const submit = () => {
    if (!canEdit) return;
    const keepNoProof = noProofIndeterminate ? { preserveNoProof: true } : {};
    if (selectedKey === MIXED_LANGUAGE_KEY) {
      // The language was left as "(mixed)": only the checking switch
      // changes, and every run keeps its own language.
      if (
        !onSubmit(
          { val: null, eastAsia: null, bidi: null, noProof },
          { preserveLanguages: true, ...keepNoProof },
        )
      ) {
        setError(copy.stale);
        return;
      }
      onClose();
      return;
    }
    const current = selectedKey === CURRENT_LANGUAGE_KEY && initialValue ? initialValue : null;
    const option =
      LANGUAGE_OPTIONS.find((candidate) => candidate.key === selectedKey) ??
      LANGUAGE_OPTIONS[0];
    const value = {
      val: current ? current.val : option.val,
      eastAsia: current ? current.eastAsia : option.eastAsia,
      bidi: current ? current.bidi : option.bidi,
      noProof,
    };
    if (!(noProofIndeterminate ? onSubmit(value, keepNoProof) : onSubmit(value))) {
      setError(copy.stale);
      return;
    }
    onClose();
  };

  return (
    <OfficeDialogShell
      open={open}
      title={copy.title}
      description={copy.description}
      defaultActionLabel={copy.apply}
      cancelActionLabel={copy.cancel}
      defaultActionDisabled={!canEdit}
      onDefaultAction={submit}
      onCancel={onClose}
      initialFocusRef={selectRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="document-proofing-language"
    >
      <label className={parts.field}>
        <span className={parts.label}>{copy.language}</span>
        <select
          disabled={!canEdit}
          ref={selectRef}
          value={selectedKey}
          data-testid="document-proofing-language-select"
          onChange={(event) => {
            setSelectedKey(event.target.value);
            setError(null);
          }}
        >
          {mixed ? (
            <option value={MIXED_LANGUAGE_KEY} disabled>
              {copy.mixed}
            </option>
          ) : null}
          {unlisted && initialValue ? (
            <option value={CURRENT_LANGUAGE_KEY} data-testid="document-proofing-language-current">
              {currentLanguageLabel(initialValue)}
            </option>
          ) : null}
          {LANGUAGE_OPTIONS.map((option) => (
            <option key={option.key} value={option.key}>
              {copy.labels[option.key]}
            </option>
          ))}
        </select>
      </label>
      <label className={parts.check}>
        <input
          disabled={!canEdit}
          type="checkbox"
          ref={noProofRef}
          checked={noProof}
          aria-checked={noProofIndeterminate ? "mixed" : noProof}
          data-testid="document-proofing-language-no-proof"
          onChange={(event) => {
            setNoProof(event.target.checked);
            setNoProofTouched(true);
            setError(null);
          }}
        />
        <span>{copy.noProof}</span>
      </label>
      {error ? (
        <p role="alert" className={parts.error}>
          {error}
        </p>
      ) : null}
    </OfficeDialogShell>
  );
}
