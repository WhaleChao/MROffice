"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useMemo, useState, type RefObject } from "react";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  previewDataQuery,
  type DataQueryDefinition,
  type DataQueryStep,
  type DataQueryTable,
  type DataQueryRun,
  type DataQueryStatus,
} from "@mrnote/core/web/lib/spreadsheet/data-queries.ts";
import { SpreadsheetCollapseDialogButton } from "./SpreadsheetRangePicker";
import type { SpreadsheetRangePick } from "./useSpreadsheetRangePick";
import { SpreadsheetDialogIcon } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

interface Props {
  translate: (key: string, values?: Record<string, string | number>) => string;
  source: DataQueryTable;
  initialDefinition: DataQueryDefinition;
  initialDestination: string;
  fixedDestination?: boolean;
  hasHeaders?: boolean;
  onHeadersChange?: (hasHeaders: boolean) => void;
  busy: boolean;
  error?: string;
  returnFocusRef: RefObject<HTMLElement | null>;
  onCancel: () => void;
  onLoad: (definition: DataQueryDefinition, destination: string) => void;
  /** The window's RefEdit: the result's first cell can be picked on the sheet. */
  rangePick?: SpreadsheetRangePick;
}

export function SpreadsheetDataQueryDialog(props: Props) {
  const { translate } = props;
  const [definition, setDefinition] = useState(props.initialDefinition);
  const [destination, setDestination] = useState(props.initialDestination);
  const [operation, setOperation] =
    useState<DataQueryStep["kind"]>("filterEquals");
  const [column, setColumn] = useState(props.source.columns[0] ?? "");
  const [value, setValue] = useState("");
  const [valueKind, setValueKind] = useState("text");
  const [selection, setSelection] = useState({
    source: props.source,
    columns: props.source.columns,
  });
  const selected =
    selection.source === props.source
      ? selection.columns
      : props.source.columns;
  const preview = useMemo(() => {
    try {
      return { table: previewDataQuery(definition, props.source), error: "" };
    } catch {
      return {
        table: null,
        error: translate("spreadsheet.helperCopy.dataQuery.aStepRefersToAMissingColumnOrUnsupportedData"),
      };
    }
  }, [definition, props.source, translate]);
  const columns = preview.table?.columns ?? props.source.columns;
  const chosenColumn = columns.includes(column) ? column : columns[0];
  const chosenColumns = selected.filter((name) => columns.includes(name));
  const addStep = () => {
    let step: DataQueryStep;
    if (operation === "selectColumns" || operation === "distinct") {
      if (!chosenColumns.length) return;
      step = { kind: operation, columns: chosenColumns };
    } else if (operation === "renameColumn") {
      if (!value.trim()) return;
      step = { kind: operation, from: chosenColumn, to: value.trim() };
    } else if (operation === "removeBlankRows") step = { kind: operation };
    else {
      if (
        valueKind === "number" &&
        (!value.trim() || !Number.isFinite(Number(value)))
      )
        return;
      step = {
        kind: "filterEquals",
        column: chosenColumn,
        value:
          valueKind === "blank"
            ? null
            : valueKind === "number"
              ? Number(value)
              : value,
      };
    }
    setDefinition((current) => ({
      ...current,
      steps: [...current.steps, step],
    }));
  };
  const labels: Record<DataQueryStep["kind"], string> = {
    selectColumns: translate("spreadsheet.helperCopy.dataQuery.chooseColumns"),
    renameColumn: translate("spreadsheet.helperCopy.dataQuery.renameColumn"),
    filterEquals: translate("spreadsheet.helperCopy.dataQuery.filterEquals"),
    removeBlankRows: translate("spreadsheet.helperCopy.dataQuery.removeBlankRows"),
    distinct: translate("spreadsheet.helperCopy.dataQuery.removeDuplicates"),
  };
  // Collapsed to the RefEdit bar while its destination is picked.
  if (props.rangePick?.picking("spreadsheet-data-query-destination")) return null;
  return (
    <OfficeDialogShell
      open
      title={translate("spreadsheet.helperCopy.dataQuery.queryEditor")}
      description={
        translate("spreadsheet.helperCopy.dataQuery.previewDoesNotChangeCellsAfterLoadingUseRefreshTo")
      }
      defaultActionLabel={translate("spreadsheet.helperCopy.dataQuery.load")}
      cancelActionLabel={translate("spreadsheet.helperCopy.dataQuery.cancel")}
      defaultActionDisabled={
        !preview.table ||
        definition.steps.length > 32 ||
        !definition.name.trim() ||
        !destination.trim()
      }
      busy={props.busy}
      onDefaultAction={() => props.onLoad(definition, destination)}
      onCancel={props.onCancel}
      returnFocusRef={props.returnFocusRef}
      testIdPrefix="spreadsheet-data-query"
      defaultActionTestId="spreadsheet-data-query-load"
    >
      <div className="grid gap-3" data-testid="spreadsheet-data-query-editor">
        {props.hasHeaders !== undefined && (
          <label className={officeDialogParts.check}>
            <input
              type="checkbox"
              checked={props.hasHeaders}
              disabled={props.busy || definition.steps.length > 0}
              data-testid="spreadsheet-data-query-headers"
              onChange={(event) =>
                props.onHeadersChange?.(event.target.checked)
              }
            />
            {translate("spreadsheet.helperCopy.dataQuery.firstRowContainsHeadersUniqueNonemptyText")}
          </label>
        )}
        <div className="grid grid-cols-2 gap-3">
          <label className={officeDialogParts.field}>
            <span className={officeDialogParts.label}>{translate("spreadsheet.helperCopy.dataQuery.queryName")}</span>
            <input
              value={definition.name}
              maxLength={200}
              data-testid="spreadsheet-data-query-name"
              onChange={(event) =>
                setDefinition({ ...definition, name: event.target.value })
              }
            />
          </label>
          <label className={officeDialogParts.field}>
            <span className={officeDialogParts.label}>{translate("spreadsheet.helperCopy.dataQuery.resultStartsAt")}</span>
            <span style={{ display: "flex", gap: 4 }}>
              <input
                value={destination}
                disabled={props.fixedDestination}
                maxLength={20}
                data-testid="spreadsheet-data-query-destination"
                onChange={(event) => setDestination(event.target.value)}
                style={{ flex: 1, minWidth: 0 }}
              />
              {props.rangePick ? (
                <SpreadsheetCollapseDialogButton
                  disabled={props.fixedDestination}
                  label={translate("spreadsheet.cfDialog.collapse")}
                  onClick={() =>
                    props.rangePick?.start({
                      id: "spreadsheet-data-query-destination",
                      dialogLabel: translate("spreadsheet.helperCopy.dataQuery.queryEditor"),
                      inputLabel: translate("spreadsheet.helperCopy.dataQuery.resultStartsAt"),
                      style: "absoluteFormula",
                      value: destination,
                      onChange: setDestination,
                    })
                  }
                  testId="spreadsheet-data-query-destination-pick"
                />
              ) : null}
            </span>
          </label>
        </div>
        <div className="flex flex-wrap items-end gap-2">
          <label className={officeDialogParts.field}>
            <span className={officeDialogParts.label}>{translate("spreadsheet.helperCopy.dataQuery.transform")}</span>
            <select
              value={operation}
              data-testid="spreadsheet-data-query-operation"
              onChange={(event) =>
                setOperation(event.target.value as DataQueryStep["kind"])
              }
            >
              {Object.entries(labels).map(([key, label]) => (
                <option key={key} value={key}>
                  {label}
                </option>
              ))}
            </select>
          </label>
          {(operation === "renameColumn" || operation === "filterEquals") && (
            <>
              <label className={officeDialogParts.field}>
                <span className={officeDialogParts.label}>{translate("spreadsheet.helperCopy.dataQuery.column")}</span>
                <select
                  value={chosenColumn}
                  onChange={(event) => setColumn(event.target.value)}
                >
                  {columns.map((name) => (
                    <option key={name}>{name}</option>
                  ))}
                </select>
              </label>
              {operation === "filterEquals" && (
                <label className={officeDialogParts.field}>
                  <span className={officeDialogParts.label}>{translate("spreadsheet.helperCopy.dataQuery.type")}</span>
                  <select
                    value={valueKind}
                    onChange={(event) => setValueKind(event.target.value)}
                  >
                    <option value="text">{translate("spreadsheet.helperCopy.dataQuery.text")}</option>
                    <option value="number">{translate("spreadsheet.helperCopy.dataQuery.number")}</option>
                    <option value="blank">{translate("spreadsheet.helperCopy.dataQuery.blank")}</option>
                  </select>
                </label>
              )}
              <label className={officeDialogParts.field}>
                <span className={officeDialogParts.label}>
                  {operation === "renameColumn"
                    ? translate("spreadsheet.helperCopy.dataQuery.newName")
                    : translate("spreadsheet.helperCopy.dataQuery.value")}
                </span>
                <input
                  data-testid="spreadsheet-data-query-value"
                  value={value}
                  maxLength={32767}
                  disabled={
                    operation === "filterEquals" && valueKind === "blank"
                  }
                  onChange={(event) => setValue(event.target.value)}
                />
              </label>
            </>
          )}
          <button
            type="button"
            className={officeDialogParts.compactButton}
            onClick={addStep}
            disabled={!preview.table || definition.steps.length >= 32}
            data-testid="spreadsheet-data-query-add-step"
          >
            <SpreadsheetDialogIcon name="addStep" />
            {translate("spreadsheet.helperCopy.dataQuery.addStep")}
          </button>
        </div>
        {(operation === "selectColumns" || operation === "distinct") && (
          <div className="flex flex-wrap gap-3">
            {columns.map((name) => (
              <label key={name} className={officeDialogParts.check}>
                <input
                  type="checkbox"
                  checked={selected.includes(name)}
                  onChange={(event) =>
                    setSelection({
                      source: props.source,
                      columns: event.target.checked
                        ? [...selected, name]
                        : selected.filter((item) => item !== name),
                    })
                  }
                />
                {name}
              </label>
            ))}
          </div>
        )}
        {definition.steps.length > 0 && (
          <ol
            className={`${officeDialogParts.list} max-h-28`}
            aria-label={translate("spreadsheet.helperCopy.dataQuery.appliedSteps")}
          >
            {definition.steps.map((step, index) => (
              <li
                key={index}
                className={`${officeDialogParts.listItem} flex items-center justify-between gap-2`}
              >
                <span>
                  {index + 1}. {labels[step.kind]}
                  {"column" in step
                    ? ` · ${step.column}`
                    : "from" in step
                      ? ` · ${step.from} → ${step.to}`
                      : "columns" in step
                        ? ` · ${step.columns.join(", ")}`
                        : ""}
                </span>
                <button
                  type="button"
                  className={officeDialogParts.iconButton}
                  aria-label={translate("spreadsheet.helperCopy.dataQuery.removeStep", { index: index + 1 })}
                  title={translate("spreadsheet.helperCopy.dataQuery.removeStep", { index: index + 1 })}
                  onClick={() =>
                    setDefinition({
                      ...definition,
                      steps: definition.steps.filter(
                        (_, item) => item !== index,
                      ),
                    })
                  }
                >
                  ×
                </button>
              </li>
            ))}
          </ol>
        )}
        {(props.error || preview.error) && (
          <p role="alert" className={officeDialogParts.error}>
            {props.error || preview.error}
          </p>
        )}
        {preview.table && (
          <>
            <div className={`${officeDialogParts.well} max-h-56 overflow-auto`}>
              <table
                className={`${officeDialogParts.table} w-full text-left`}
                data-testid="spreadsheet-data-query-preview"
              >
                <thead>
                  <tr>
                    {preview.table.columns.map((name) => (
                      <th key={name}>
                        {name}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.table.rows.slice(0, 50).map((row, index) => (
                    <tr key={index}>
                      {row.map((cell, col) => (
                        <td key={col}>
                          {cell === null ? "" : String(cell)}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className={officeDialogParts.hint}>
              {translate("spreadsheet.helperCopy.dataQuery.rowsPreviewShowsTheFirst50", { count: preview.table.rows.length })}
            </p>
          </>
        )}
      </div>
    </OfficeDialogShell>
  );
}

export function SpreadsheetQueryConnectionsDialog(props: {
  translate: (key: string, values?: Record<string, string | number>) => string;
  runs: DataQueryRun[];
  busy: boolean;
  error?: string;
  onClose: () => void;
  onEdit: (run: DataQueryRun) => void;
  onRefresh: (run: DataQueryRun) => void;
  /** Query Properties of one query (Excel: the pane's Properties). */
  onProperties?: (run: DataQueryRun) => void;
  /** The query's description, shown under its name as Excel's pane does. */
  descriptionFor?: (run: DataQueryRun) => string;
  statusFor: (run: DataQueryRun) => Promise<DataQueryStatus>;
}) {
  const [statuses, setStatuses] = useState<Record<string, DataQueryStatus>>({});
  const { runs, statusFor, translate } = props;
  useEffect(() => {
    let live = true;
    void Promise.all(
      runs.map(async (run) => [run.queryId, await statusFor(run)] as const),
    ).then((entries) => {
      if (live) setStatuses(Object.fromEntries(entries));
    });
    return () => {
      live = false;
    };
  }, [runs, statusFor]);
  const statusLabels: Record<DataQueryStatus, string> = {
    current: translate("spreadsheet.helperCopy.dataQuery.matchesCurrentSource"),
    stale: translate("spreadsheet.helperCopy.dataQuery.sourceChanged"),
    edited: translate("spreadsheet.helperCopy.dataQuery.resultEdited"),
    blocked: translate("spreadsheet.helperCopy.dataQuery.cannotRefresh"),
  };
  return (
    <OfficeDialogShell
      open
      title={translate("spreadsheet.helperCopy.dataQuery.queriesConnections")}
      description={
        translate("spreadsheet.helperCopy.dataQuery.localQueriesInThisWorkbook")
      }
      defaultActionLabel={translate("spreadsheet.helperCopy.dataQuery.close")}
      cancelActionLabel={translate("spreadsheet.helperCopy.dataQuery.cancel")}
      onDefaultAction={props.onClose}
      onCancel={props.onClose}
      testIdPrefix="spreadsheet-query-connections"
    >
      {props.error && (
        <p role="alert" className={officeDialogParts.error}>
          {props.error}
        </p>
      )}
      {props.runs.length === 0 ? (
        <p className={officeDialogParts.hint} data-testid="spreadsheet-query-connections-empty">
          {translate("spreadsheet.data.noQueries")}
        </p>
      ) : null}
      <ul className={officeDialogParts.list}>
        {props.runs.map((run) => (
          <li
            key={run.queryId}
            className={`${officeDialogParts.listItem} flex items-center gap-2`}
          >
            <span className="flex-1">
              {run.definition.name}
              {props.descriptionFor?.(run) ? (
                <span
                  className={`${officeDialogParts.hint} block`}
                  data-testid={`spreadsheet-query-description-${run.queryId}`}
                >
                  {props.descriptionFor(run)}
                </span>
              ) : null}
              <span className={`${officeDialogParts.hint} block`}>
                {run.definition.sourceName} → {run.targetName}
              </span>
              {statuses[run.queryId] && (
                <span
                  className={`${officeDialogParts.hint} block`}
                  data-testid={`spreadsheet-query-status-${run.queryId}`}
                >
                  {statusLabels[statuses[run.queryId]]}
                </span>
              )}
            </span>
            <button
              type="button"
              className={officeDialogParts.compactButton}
              disabled={props.busy}
              onClick={() => props.onEdit(run)}
            >
              {translate("spreadsheet.helperCopy.dataQuery.edit")}
            </button>
            <button
              type="button"
              className={officeDialogParts.compactButton}
              disabled={props.busy}
              onClick={() => props.onRefresh(run)}
            >
              {translate("spreadsheet.helperCopy.dataQuery.refresh")}
            </button>
            {props.onProperties ? (
              <button
                type="button"
                className={officeDialogParts.compactButton}
                data-testid={`spreadsheet-query-properties-open-${run.queryId}`}
                onClick={() => props.onProperties?.(run)}
              >
                {translate("spreadsheet.helperCopy.queryProperties.properties")}
              </button>
            ) : null}
          </li>
        ))}
      </ul>
    </OfficeDialogShell>
  );
}
