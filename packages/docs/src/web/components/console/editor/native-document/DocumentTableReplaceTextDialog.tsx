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
import {
  DOCUMENT_TABLE_REPLACE_TEXT_MAX_LENGTH,
  type DocumentTableReplaceTextResult,
} from "./documentTableReplaceText";

export interface DocumentTableReplaceTextDialogProps {
  readonly canEdit?: boolean;
  readonly open: boolean;
  readonly locale: string;
  readonly returnFocusRef: RefObject<HTMLElement | null>;
  readonly onSubmit: (
    query: string,
    replacement: string,
    options: { readonly caseSensitive: boolean; readonly wholeWord: boolean },
  ) => DocumentTableReplaceTextResult;
  readonly onClose: () => void;
}

export function DocumentTableReplaceTextDialog({
  canEdit = true,
  open,
  returnFocusRef,
  onSubmit,
  onClose,
}: DocumentTableReplaceTextDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const copy = useMemo(
    () => ({
      title: t("toolbar.tableReplaceTextDialog.title"),
      description: t("toolbar.tableReplaceTextDialog.description"),
      find: t("toolbar.tableReplaceTextDialog.find"),
      replace: t("toolbar.tableReplaceTextDialog.replace"),
      caseSensitive: t("toolbar.tableReplaceTextDialog.caseSensitive"),
      wholeWord: t("toolbar.tableReplaceTextDialog.wholeWord"),
      apply: t("toolbar.tableReplaceTextDialog.apply"),
      cancel: t("toolbar.tableReplaceTextDialog.cancel"),
      unchanged: t("toolbar.tableReplaceTextDialog.unchanged"),
      stale: t("toolbar.tableReplaceTextDialog.stale"),
      invalid: t("toolbar.tableReplaceTextDialog.invalid"),
    }),
    [t],
  );
  const [query, setQuery] = useState("");
  const [replacement, setReplacement] = useState("");
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [wholeWord, setWholeWord] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const queryRef = useRef<HTMLInputElement | null>(null);
  const valid =
    query.length > 0 &&
    query.length <= DOCUMENT_TABLE_REPLACE_TEXT_MAX_LENGTH &&
    replacement.length <= DOCUMENT_TABLE_REPLACE_TEXT_MAX_LENGTH;

  const submit = useCallback(() => {
    if (!canEdit) return;
    if (!valid) {
      setError(copy.invalid);
      return;
    }
    const result = onSubmit(query, replacement, {
      caseSensitive,
      wholeWord,
    });
    if (result === "applied") {
      onClose();
      return;
    }
    setError(
      result === "unchanged"
        ? copy.unchanged
        : result === "stale"
          ? copy.stale
          : copy.invalid,
    );
  }, [
    canEdit,
    caseSensitive,
    copy.invalid,
    copy.stale,
    copy.unchanged,
    onClose,
    onSubmit,
    query,
    replacement,
    valid,
    wholeWord,
  ]);

  return (
    <OfficeDialogShell
      open={open}
      title={copy.title}
      description={copy.description}
      defaultActionLabel={copy.apply}
      cancelActionLabel={copy.cancel}
      defaultActionDisabled={!canEdit || !valid}
      onDefaultAction={submit}
      onCancel={onClose}
      initialFocusRef={queryRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="document-table-replace-text"
    >
      <label className={parts.field}>
        <span className={parts.label}>{copy.find}</span>
        <input
          disabled={!canEdit}
          ref={queryRef}
          type="search"
          value={query}
          maxLength={DOCUMENT_TABLE_REPLACE_TEXT_MAX_LENGTH}
          data-testid="document-table-replace-text-query"
          onChange={(event) => {
            setQuery(event.target.value);
            setError(null);
          }}
        />
      </label>
      <label className={parts.field}>
        <span className={parts.label}>{copy.replace}</span>
        <input
          disabled={!canEdit}
          type="text"
          value={replacement}
          maxLength={DOCUMENT_TABLE_REPLACE_TEXT_MAX_LENGTH}
          data-testid="document-table-replace-text-replacement"
          onChange={(event) => {
            setReplacement(event.target.value);
            setError(null);
          }}
        />
      </label>
      <label className={parts.check}>
        <input
          disabled={!canEdit}
          type="checkbox"
          checked={caseSensitive}
          data-testid="document-table-replace-text-case-sensitive"
          onChange={(event) => {
            setCaseSensitive(event.target.checked);
            setError(null);
          }}
        />
        <span>{copy.caseSensitive}</span>
      </label>
      <label className={parts.check}>
        <input
          disabled={!canEdit}
          type="checkbox"
          checked={wholeWord}
          data-testid="document-table-replace-text-whole-word"
          onChange={(event) => {
            setWholeWord(event.target.checked);
            setError(null);
          }}
        />
        <span>{copy.wholeWord}</span>
      </label>
      {error ? (
        <p
          role="alert"
          className={parts.error}
          data-testid="document-table-replace-text-error"
        >
          {error}
        </p>
      ) : null}
    </OfficeDialogShell>
  );
}
