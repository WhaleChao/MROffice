"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useMemo, useRef, useState, type RefObject } from "react";
import { Minus, Plus } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import {
  officeDialogParts,
  useOfficeDialogChrome,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  MAX_INSERT_TABLE_CELLS,
  MAX_INSERT_TABLE_COLUMNS,
  MAX_INSERT_TABLE_ROWS,
  MIN_INSERT_TABLE_COLUMNS,
  MIN_INSERT_TABLE_ROWS,
  type TableInsertDimensionError,
  validateTableInsertDimensions,
} from "../presentation-table-insert";

export interface PresentationTableInsertDialogProps {
  readonly open: boolean;
  readonly initialRows: number;
  readonly initialColumns: number;
  readonly returnFocusRef: RefObject<HTMLElement | null>;
  readonly onInsert: (rows: number, columns: number) => boolean;
  readonly onClose: () => void;
}

function stepDimension(
  draft: string,
  delta: -1 | 1,
  minimum: number,
  maximum: number,
  fallback: number,
): string {
  const trimmed = draft.trim();
  const current = /^[0-9]+$/.test(trimmed) ? Number(trimmed) : fallback;
  const next = Math.min(maximum, Math.max(minimum, current + delta));
  return String(next);
}

interface DimensionFieldProps {
  readonly kind: "rows" | "columns";
  readonly label: string;
  readonly value: string;
  readonly error: string | null;
  readonly decrementLabel: string;
  readonly incrementLabel: string;
  readonly inputRef?: RefObject<HTMLInputElement | null>;
  readonly onChange: (value: string) => void;
}

function DimensionField({
  kind,
  label,
  value,
  error,
  decrementLabel,
  incrementLabel,
  inputRef,
  onChange,
}: DimensionFieldProps) {
  const rows = kind === "rows";
  const minimum = rows ? MIN_INSERT_TABLE_ROWS : MIN_INSERT_TABLE_COLUMNS;
  const maximum = rows ? MAX_INSERT_TABLE_ROWS : MAX_INSERT_TABLE_COLUMNS;
  const errorId = error ? `presentation-table-insert-${kind}-error` : undefined;
  const decrementDisabled = value === String(minimum);
  const incrementDisabled = value === String(maximum);
  // Under the Office dialog chrome the value and its step buttons are one
  // stepper field; elsewhere the standalone framed row below.
  const officeChrome = useOfficeDialogChrome() !== null;
  const stepButtonClass = officeChrome
    ? officeDialogParts.stepperButton
    : undefined;

  return (
    <label
      className={
        officeChrome
          ? officeDialogParts.field
          : "grid gap-1.5 text-sm font-medium text-foreground"
      }
    >
      <span>{label}</span>
      <span
        className={
          officeChrome
            ? officeDialogParts.stepper
            : "grid grid-cols-[2.25rem_minmax(0,1fr)_2.25rem] overflow-hidden rounded-md border border-input bg-background focus-within:ring-2 focus-within:ring-ring focus-within:ring-offset-2"
        }
      >
        <button
          type="button"
          className={
            stepButtonClass ??
            "inline-flex h-9 items-center justify-center border-r border-input text-foreground transition-colors hover:bg-accent focus-visible:z-10 focus-visible:outline-none disabled:cursor-not-allowed disabled:opacity-45"
          }
          aria-label={decrementLabel}
          title={decrementLabel}
          disabled={decrementDisabled}
          data-testid={`presentation-table-insert-${kind}-decrement`}
          onClick={() =>
            onChange(stepDimension(value, -1, minimum, maximum, 2))
          }
        >
          <Minus size={14} aria-hidden="true" />
        </button>
        <input
          ref={inputRef}
          type="text"
          inputMode="numeric"
          pattern="[0-9]*"
          autoComplete="off"
          value={value}
          aria-label={label}
          aria-invalid={error ? "true" : undefined}
          aria-describedby={errorId}
          data-testid={`presentation-table-insert-${kind}`}
          className={
            officeChrome
              ? "tabular-nums"
              : "h-9 min-w-0 border-0 bg-transparent px-3 text-center text-sm tabular-nums outline-none"
          }
          onChange={(event) => onChange(event.target.value)}
        />
        <button
          type="button"
          className={
            stepButtonClass ??
            "inline-flex h-9 items-center justify-center border-l border-input text-foreground transition-colors hover:bg-accent focus-visible:z-10 focus-visible:outline-none disabled:cursor-not-allowed disabled:opacity-45"
          }
          aria-label={incrementLabel}
          title={incrementLabel}
          disabled={incrementDisabled}
          data-testid={`presentation-table-insert-${kind}-increment`}
          onClick={() =>
            onChange(stepDimension(value, 1, minimum, maximum, 2))
          }
        >
          <Plus size={14} aria-hidden="true" />
        </button>
      </span>
      {error ? (
        <span id={errorId} className={officeDialogParts.error} role="alert">
          {error}
        </span>
      ) : null}
    </label>
  );
}

/** Office-style bounded row/column dialog behind Insert > Table. */
export function PresentationTableInsertDialog({
  open,
  initialRows,
  initialColumns,
  returnFocusRef,
  onInsert,
  onClose,
}: PresentationTableInsertDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const columnsInputRef = useRef<HTMLInputElement | null>(null);
  const [rowsDraft, setRowsDraft] = useState(String(initialRows));
  const [columnsDraft, setColumnsDraft] = useState(String(initialColumns));
  const [commandError, setCommandError] = useState<string | null>(null);

  const validation = useMemo(
    () => validateTableInsertDimensions(rowsDraft, columnsDraft),
    [columnsDraft, rowsDraft],
  );
  const fieldErrorMessage = (
    error: TableInsertDimensionError | null,
    kind: "rows" | "columns",
  ): string | null => {
    if (error === "required") {
      return t("presentation.tableInsert.required");
    }
    if (error === "integer") {
      return t("presentation.tableInsert.integer");
    }
    if (error === "range") {
      return kind === "rows"
        ? t("presentation.tableInsert.rowsRange", {
            minimum: MIN_INSERT_TABLE_ROWS,
            maximum: MAX_INSERT_TABLE_ROWS,
          })
        : t("presentation.tableInsert.columnsRange", {
            minimum: MIN_INSERT_TABLE_COLUMNS,
            maximum: MAX_INSERT_TABLE_COLUMNS,
          });
    }
    return null;
  };
  const rowError = fieldErrorMessage(validation.rowError, "rows");
  const columnError = fieldErrorMessage(validation.columnError, "columns");

  const submit = () => {
    if (
      !validation.ok ||
      validation.rows === null ||
      validation.columns === null
    ) {
      return;
    }
    if (!onInsert(validation.rows, validation.columns)) {
      setCommandError(t("presentation.tableInsert.commandError"));
      return;
    }
    onClose();
  };

  const close = () => {
    setCommandError(null);
    onClose();
  };

  return (
    <OfficeDialogShell
      open={open}
      title={t("toolbar.insertTableSizeTitle")}
      description={t("presentation.tableInsert.description")}
      defaultActionLabel={t("toolbar.insertTableSizeApply")}
      cancelActionLabel={t("toolbar.cancel")}
      onDefaultAction={submit}
      onCancel={close}
      defaultActionDisabled={!validation.ok}
      initialFocusRef={columnsInputRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="presentation-table-insert"
      className="sm:max-w-[420px]"
    >
      <div className="grid gap-4">
        <div className="grid grid-cols-2 gap-3">
          <DimensionField
            kind="columns"
            label={t("toolbar.insertTableColumns")}
            value={columnsDraft}
            error={columnError}
            decrementLabel={t("presentation.tableInsert.decrement", {
              label: t("toolbar.insertTableColumns"),
            })}
            incrementLabel={t("presentation.tableInsert.increment", {
              label: t("toolbar.insertTableColumns"),
            })}
            inputRef={columnsInputRef}
            onChange={(value) => {
              setColumnsDraft(value);
              setCommandError(null);
            }}
          />
          <DimensionField
            kind="rows"
            label={t("toolbar.insertTableRows")}
            value={rowsDraft}
            error={rowError}
            decrementLabel={t("presentation.tableInsert.decrement", {
              label: t("toolbar.insertTableRows"),
            })}
            incrementLabel={t("presentation.tableInsert.increment", {
              label: t("toolbar.insertTableRows"),
            })}
            onChange={(value) => {
              setRowsDraft(value);
              setCommandError(null);
            }}
          />
        </div>
        {validation.capacityExceeded ? (
          <p
            className={officeDialogParts.message}
            data-tone="error"
            role="alert"
            data-testid="presentation-table-insert-capacity-error"
          >
            {t("presentation.tableInsert.capacity", {
              maximum: MAX_INSERT_TABLE_CELLS,
            })}
          </p>
        ) : null}
        {commandError ? (
          <p
            className={officeDialogParts.message}
            data-tone="error"
            role="alert"
            data-testid="presentation-table-insert-command-error"
          >
            {commandError}
          </p>
        ) : null}
      </div>
    </OfficeDialogShell>
  );
}
