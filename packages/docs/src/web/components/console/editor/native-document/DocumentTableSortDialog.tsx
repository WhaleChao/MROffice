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
import { DocumentDialogIcon } from "@mrnote/core/web/components/console/editor/native-document/DocumentDialogIcon.tsx";
import type {
  SortDirection,
  SortValueType,
  TableSortCriterion,
} from "./documentTextTransforms";

export interface DocumentTableSortSettings {
  readonly criteria: readonly TableSortCriterion[];
  readonly caseSensitive: boolean;
  readonly hasHeader: boolean;
}

export interface DocumentTableSortDialogProps {
  readonly canEdit?: boolean;
  readonly open: boolean;
  readonly locale: string;
  readonly columnLabels: readonly string[];
  readonly initialColumnIndex: number;
  readonly initialHasHeader: boolean;
  readonly returnFocusRef: RefObject<HTMLElement | null>;
  readonly onSubmit: (settings: DocumentTableSortSettings) => boolean;
  readonly onClose: () => void;
}

interface SortCopy {
  readonly title: string;
  readonly description: string;
  readonly criterion: readonly [string, string, string];
  readonly none: string;
  readonly column: (index: number, name: string) => string;
  readonly type: string;
  readonly text: string;
  readonly number: string;
  readonly date: string;
  readonly dateHelp: string;
  readonly ascending: string;
  readonly descending: string;
  readonly headerRow: string;
  readonly caseSensitive: string;
  readonly sort: string;
  readonly cancel: string;
  readonly changed: string;
}

const TABLE_SORT_TYPE_ICON = {
  text: "text.type",
  number: "symbol.hash",
  date: "calendar.date",
} as const;

function sortCopy(
  t: (key: string, values?: Record<string, string | number>) => string,
): SortCopy {
  const k = (leaf: string) => t(`toolbar.tableSortDialog.${leaf}`);
  return {
    title: k("title"),
    description: k("description"),
    criterion: [k("sortBy"), k("thenBy"), k("thenByThird")],
    none: k("none"),
    column: (index, name) =>
      name
        ? t("toolbar.tableSortDialog.columnNamed", { index, name })
        : t("toolbar.tableSortDialog.column", { index }),
    type: k("type"),
    text: k("text"),
    number: k("number"),
    date: k("date"),
    dateHelp: k("dateHelp"),
    ascending: k("ascending"),
    descending: k("descending"),
    headerRow: k("headerRow"),
    caseSensitive: k("caseSensitive"),
    sort: k("sort"),
    cancel: k("cancel"),
    changed: k("changed"),
  };
}

interface SortCriterionDraft {
  readonly columnIndex: number | null;
  readonly direction: SortDirection;
  readonly sortType: SortValueType;
}

export function DocumentTableSortDialog({
  canEdit = true,
  open,
  columnLabels,
  initialColumnIndex,
  initialHasHeader,
  returnFocusRef,
  onSubmit,
  onClose,
}: DocumentTableSortDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const copy = useMemo(() => sortCopy(t), [t]);
  const columnRef = useRef<HTMLSelectElement | null>(null);
  const safeInitialColumn = Math.min(
    Math.max(Math.trunc(initialColumnIndex) || 0, 0),
    Math.max(0, columnLabels.length - 1),
  );
  const [criteria, setCriteria] = useState<readonly SortCriterionDraft[]>([
    { columnIndex: safeInitialColumn, direction: "asc", sortType: "text" },
    { columnIndex: null, direction: "asc", sortType: "text" },
    { columnIndex: null, direction: "asc", sortType: "text" },
  ]);
  const [hasHeader, setHasHeader] = useState(initialHasHeader);
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const updateCriterion = useCallback(
    (index: number, patch: Partial<SortCriterionDraft>) => {
      setCriteria((current) =>
        current.map((criterion, criterionIndex) =>
          criterionIndex === index ? { ...criterion, ...patch } : criterion,
        ),
      );
      setError(null);
    },
    [],
  );

  const updateCriterionColumn = useCallback(
    (index: number, columnIndex: number | null) => {
      setCriteria((current) =>
        current.map((criterion, criterionIndex) => {
          if (criterionIndex === index) return { ...criterion, columnIndex };
          return columnIndex === null && criterionIndex > index
            ? { ...criterion, columnIndex: null }
            : criterion;
        }),
      );
      setError(null);
    },
    [],
  );

  const hasTextCriterion = criteria.some(
    (criterion) =>
      criterion.columnIndex !== null && criterion.sortType === "text",
  );

  const submit = useCallback(() => {
    if (!canEdit) return;
    const activeCriteria = criteria
      .filter(
        (criterion): criterion is SortCriterionDraft & { columnIndex: number } =>
          criterion.columnIndex !== null,
      )
      .map(({ columnIndex, direction, sortType }) => ({
        columnIndex,
        direction,
        sortType,
      }));
    if (
      !onSubmit({
        criteria: activeCriteria,
        caseSensitive: hasTextCriterion && caseSensitive,
        hasHeader,
      })
    ) {
      setError(copy.changed);
      return;
    }
    onClose();
  }, [
    canEdit,
    caseSensitive,
    copy.changed,
    criteria,
    hasHeader,
    hasTextCriterion,
    onClose,
    onSubmit,
  ]);

  return (
    <OfficeDialogShell
      open={open}
      title={copy.title}
      description={copy.description}
      defaultActionLabel={copy.sort}
      cancelActionLabel={copy.cancel}
      defaultActionDisabled={!canEdit || columnLabels.length === 0}
      onDefaultAction={submit}
      onCancel={onClose}
      initialFocusRef={columnRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="document-table-sort"
    >
      <div className="grid max-h-[56vh] gap-3 overflow-y-auto pr-1">
        {criteria.map((criterion, index) => {
          const criterionLabel =
            copy.criterion[index] ?? copy.criterion[copy.criterion.length - 1];
          const levelEnabled =
            index === 0 || criteria[index - 1]?.columnIndex !== null;
          const controlsDisabled =
            !levelEnabled || criterion.columnIndex === null;
          return (
            <fieldset
              key={index}
              className={parts.group}
              disabled={!canEdit || !levelEnabled}
            >
              <legend className={parts.sectionTitle}>
                {criterionLabel}
              </legend>
              <select
                ref={index === 0 ? columnRef : undefined}
                value={criterion.columnIndex ?? -1}
                data-testid={
                  index === 0
                    ? "document-table-sort-column"
                    : `document-table-sort-column-${index + 1}`
                }
                aria-label={criterionLabel}
                onChange={(event) => {
                  const next = Number(event.target.value);
                  updateCriterionColumn(index, next < 0 ? null : next);
                }}
              >
                {index > 0 ? <option value={-1}>{copy.none}</option> : null}
                {columnLabels.map((label, columnIndex) => (
                  <option key={columnIndex} value={columnIndex}>
                    {copy.column(
                      columnIndex + 1,
                      label.trim().slice(0, 80),
                    )}
                  </option>
                ))}
              </select>

              <div
                className="grid grid-cols-3 gap-2"
                role="group"
                aria-label={`${criterionLabel} · ${copy.type}`}
              >
                {(
                  [
                    ["text", copy.text],
                    ["number", copy.number],
                    ["date", copy.date],
                  ] as const
                ).map(([value, label]) => (
                  <label
                    key={value}
                    className={parts.check}
                  >
                    <input
                      type="radio"
                      name={`document-table-sort-type-${index}`}
                      value={value}
                      checked={criterion.sortType === value}
                      disabled={controlsDisabled}
                      onChange={() =>
                        updateCriterion(index, { sortType: value })
                      }
                    />
                    <DocumentDialogIcon icon={TABLE_SORT_TYPE_ICON[value]} />
                    {label}
                  </label>
                ))}
              </div>

              <div
                className="grid grid-cols-2 gap-2"
                role="group"
                aria-label={`${criterionLabel} · ${copy.title}`}
              >
                {(
                  [
                    ["asc", copy.ascending],
                    ["desc", copy.descending],
                  ] as const
                ).map(([value, label]) => (
                  <label
                    key={value}
                    className={parts.check}
                  >
                    <input
                      type="radio"
                      name={`document-table-sort-direction-${index}`}
                      value={value}
                      checked={criterion.direction === value}
                      disabled={controlsDisabled}
                      onChange={() =>
                        updateCriterion(index, { direction: value })
                      }
                    />
                    <DocumentDialogIcon icon={value === "asc" ? "table.sortAscending" : "table.sortDescending"} />
                    {label}
                  </label>
                ))}
              </div>
            </fieldset>
          );
        })}
      </div>

      <p className={parts.hint}>{copy.dateHelp}</p>

      <label className={parts.check}>
        <input
          disabled={!canEdit}
          type="checkbox"
          checked={hasHeader}
          data-testid="document-table-sort-header"
          onChange={(event) => setHasHeader(event.target.checked)}
        />
        {copy.headerRow}
      </label>
      <label className={parts.check}>
        <input
          type="checkbox"
          checked={caseSensitive}
          disabled={!canEdit || !hasTextCriterion}
          data-testid="document-table-sort-case-sensitive"
          onChange={(event) => setCaseSensitive(event.target.checked)}
        />
        {copy.caseSensitive}
      </label>

      {error ? (
        <p
          role="alert"
          className={parts.error}
          data-testid="document-table-sort-error"
        >
          {error}
        </p>
      ) : null}
    </OfficeDialogShell>
  );
}
