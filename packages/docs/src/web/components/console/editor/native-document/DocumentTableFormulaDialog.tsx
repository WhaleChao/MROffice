"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type RefObject,
} from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";

import { DocumentDialogIcon } from "@mrnote/core/web/components/console/editor/native-document/DocumentDialogIcon.tsx";
import {
  TABLE_FORMULA_FUNCTIONS,
  formulaNameReadsAsCellReference,
  type TableFormulaFunction,
} from "@mrnote/core/web/components/console/editor/native-document/tableCellFormula.ts";
import {
  CELL_FORMULA_NUMBER_FORMAT_VALUES,
  MAX_CELL_FORMULA_LENGTH,
  normalizeCellFormula,
  normalizeCellFormulaNumberFormat,
  serializeCellFormulaFieldInstruction,
  type CellFormulaNumberFormat,
} from "@mrnote/core/web/components/console/editor/native-document/tableCellStyle.ts";

export interface DocumentTableFormulaSettings {
  readonly formula: string;
  readonly numberFormat: CellFormulaNumberFormat | null;
}

export interface DocumentTableFormulaDialogProps {
  readonly canEdit?: boolean;
  readonly open: boolean;
  readonly locale: string;
  readonly initialFormula: string | null;
  readonly initialNumberFormat: CellFormulaNumberFormat | null;
  readonly bookmarkNames: readonly string[];
  readonly returnFocusRef: RefObject<HTMLElement | null>;
  readonly onSubmit: (settings: DocumentTableFormulaSettings) => boolean;
  readonly onClear: () => boolean;
  readonly onClose: () => void;
}

interface FormulaCopy {
  readonly title: string;
  readonly description: string;
  readonly formula: string;
  readonly formulaPlaceholder: string;
  readonly numberFormat: string;
  readonly general: string;
  readonly pasteFunction: string;
  readonly chooseFunction: string;
  readonly pasteBookmark: string;
  readonly chooseBookmark: string;
  readonly noBookmarks: string;
  readonly invalidFormula: string;
  readonly selectionChanged: string;
  readonly clear: string;
  readonly apply: string;
  readonly cancel: string;
}

function formulaCopy(t: (key: string) => string): FormulaCopy {
  const k = (leaf: string) => t(`toolbar.tableFormulaDialog.${leaf}`);
  return {
    title: k("title"),
    description: k("description"),
    formula: k("formula"),
    formulaPlaceholder: k("formulaPlaceholder"),
    numberFormat: k("numberFormat"),
    general: k("general"),
    pasteFunction: k("pasteFunction"),
    chooseFunction: k("chooseFunction"),
    pasteBookmark: k("pasteBookmark"),
    chooseBookmark: k("chooseBookmark"),
    noBookmarks: k("noBookmarks"),
    invalidFormula: k("invalidFormula"),
    selectionChanged: k("selectionChanged"),
    clear: k("clear"),
    apply: k("apply"),
    cancel: k("cancel"),
  };
}

export function DocumentTableFormulaDialog({
  canEdit = true,
  open,
  initialFormula,
  initialNumberFormat,
  bookmarkNames,
  returnFocusRef,
  onSubmit,
  onClear,
  onClose,
}: DocumentTableFormulaDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const copy = useMemo(() => formulaCopy(t), [t]);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const initialSelectionAppliedRef = useRef(false);
  const [formulaDraft, setFormulaDraft] = useState(
    initialFormula ?? "=SUM(ABOVE)",
  );
  const [numberFormat, setNumberFormat] =
    useState<CellFormulaNumberFormat | null>(initialNumberFormat);
  const [error, setError] = useState<string | null>(null);
  // A bookmark named like a cell (A1, R1C1) is read as that cell by the
  // formula, as Word reads it: it is listed but cannot be pasted, and the
  // dialog says why instead of inserting a name that means another value
  
  const cellNamedBookmarks = useMemo(
    () => bookmarkNames.filter((name) => formulaNameReadsAsCellReference(name)),
    [bookmarkNames],
  );

  useEffect(() => {
    if (!open) {
      initialSelectionAppliedRef.current = false;
      return;
    }
    if (initialSelectionAppliedRef.current) return;
    initialSelectionAppliedRef.current = true;
    const frame = requestAnimationFrame(() => {
      const input = inputRef.current;
      if (!input) return;
      input.focus();
      // Word opens the dialog with the current expression selected, preserving
      // the leading '='. Choosing Paste Function can therefore replace the
      // default/current expression in one action instead of appending junk.
      input.setSelectionRange(
        formulaDraft.startsWith("=") ? 1 : 0,
        formulaDraft.length,
      );
    });
    return () => cancelAnimationFrame(frame);
  }, [formulaDraft, open]);

  const submit = useCallback(() => {
    if (!canEdit) return;
    const formula = normalizeCellFormula(formulaDraft);
    if (
      !formula ||
      !serializeCellFormulaFieldInstruction(formula, numberFormat)
    ) {
      setError(copy.invalidFormula);
      return;
    }
    if (!onSubmit({ formula, numberFormat })) {
      setError(copy.selectionChanged);
      return;
    }
    onClose();
  }, [canEdit, copy, formulaDraft, numberFormat, onClose, onSubmit]);

  const clear = useCallback(() => {
    if (!canEdit) return;
    if (!onClear()) {
      setError(copy.selectionChanged);
      return;
    }
    onClose();
  }, [canEdit, copy.selectionChanged, onClear, onClose]);

  const pasteFunction = useCallback(
    (name: TableFormulaFunction) => {
      const input = inputRef.current;
      const start = input?.selectionStart ?? formulaDraft.length;
      const end = input?.selectionEnd ?? start;
      const insertion = `${name}()`;
      const prefix = formulaDraft.trim() ? "" : "=";
      const next = `${formulaDraft.slice(0, start)}${prefix}${insertion}${formulaDraft.slice(end)}`;
      const caret = start + prefix.length + name.length + 1;
      setFormulaDraft(next);
      setError(null);
      requestAnimationFrame(() => {
        inputRef.current?.focus();
        inputRef.current?.setSelectionRange(caret, caret);
      });
    },
    [formulaDraft],
  );

  const pasteBookmark = useCallback(
    (name: string) => {
      const input = inputRef.current;
      const start = input?.selectionStart ?? formulaDraft.length;
      const end = input?.selectionEnd ?? start;
      const prefix = formulaDraft.trim() ? "" : "=";
      const next = `${formulaDraft.slice(0, start)}${prefix}${name}${formulaDraft.slice(end)}`;
      const caret = start + prefix.length + name.length;
      setFormulaDraft(next);
      setError(null);
      requestAnimationFrame(() => {
        inputRef.current?.focus();
        inputRef.current?.setSelectionRange(caret, caret);
      });
    },
    [formulaDraft],
  );

  const changeNumberFormat = useCallback(
    (event: ChangeEvent<HTMLSelectElement>) => {
      setNumberFormat(
        event.target.value === ""
          ? null
          : normalizeCellFormulaNumberFormat(event.target.value),
      );
      setError(null);
    },
    [],
  );

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
      initialFocusRef={inputRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="document-table-formula"
    >
      <label className={parts.field}>
        <span className={parts.label}>{copy.formula}</span>
        <input
          disabled={!canEdit}
          ref={inputRef}
          type="text"
          className={parts.mono}
          data-testid="document-toolbar-cell-formula-input"
          value={formulaDraft}
          maxLength={MAX_CELL_FORMULA_LENGTH}
          placeholder={copy.formulaPlaceholder}
          onChange={(event) => {
            setFormulaDraft(event.target.value);
            setError(null);
          }}
        />
      </label>

      <div className="grid gap-3 sm:grid-cols-3">
        <label className={parts.field}>
          <span className={parts.label}>{copy.numberFormat}</span>
          <select
            disabled={!canEdit}
            data-testid="document-toolbar-cell-formula-number-format"
            value={numberFormat ?? ""}
            onChange={changeNumberFormat}
          >
            <option value="">{copy.general}</option>
            {CELL_FORMULA_NUMBER_FORMAT_VALUES.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </label>

        <label className={parts.field}>
          <span className={parts.label}>{copy.pasteFunction}</span>
          <select
            disabled={!canEdit}
            data-testid="document-toolbar-cell-formula-function"
            value=""
            onChange={(event) => {
              const selected = TABLE_FORMULA_FUNCTIONS.find(
                (value) => value === event.target.value,
              );
              if (selected) pasteFunction(selected);
            }}
          >
            <option value="" disabled>
              {copy.chooseFunction}
            </option>
            {TABLE_FORMULA_FUNCTIONS.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </label>

        <label className={parts.field}>
          <span className={parts.label}>{copy.pasteBookmark}</span>
          <select
            data-testid="document-toolbar-cell-formula-bookmark"
            value=""
            disabled={!canEdit || bookmarkNames.length === 0}
            title={bookmarkNames.length === 0 ? copy.noBookmarks : undefined}
            onChange={(event) => {
              const selected = bookmarkNames.find(
                (value) => value === event.target.value,
              );
              if (selected && !formulaNameReadsAsCellReference(selected)) {
                pasteBookmark(selected);
              }
            }}
          >
            <option value="" disabled>
              {bookmarkNames.length === 0
                ? copy.noBookmarks
                : copy.chooseBookmark}
            </option>
            {bookmarkNames.map((name) => (
              <option
                key={name}
                value={name}
                disabled={formulaNameReadsAsCellReference(name)}
              >
                {name}
              </option>
            ))}
          </select>
        </label>
      </div>

      <div className="flex items-center gap-3">
        <button
          disabled={!canEdit}
          type="button"
          className={parts.compactButton}
          data-testid="document-toolbar-cell-formula-clear"
          onClick={clear}
        >
          <DocumentDialogIcon icon="format.clear" />
          {copy.clear}
        </button>
      </div>

      {cellNamedBookmarks.length > 0 ? (
        <p
          className={parts.hint}
          data-testid="document-toolbar-cell-formula-bookmark-cell-names"
        >
          {t("toolbar.tableFormulaDialog.bookmarkReadsAsCell", {
            names: cellNamedBookmarks.join(", "),
          })}
        </p>
      ) : null}

      {error ? (
        <p className={parts.error} role="alert">
          {error}
        </p>
      ) : null}
    </OfficeDialogShell>
  );
}
