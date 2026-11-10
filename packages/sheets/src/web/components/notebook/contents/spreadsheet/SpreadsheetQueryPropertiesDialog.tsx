"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useId, useState, type KeyboardEvent, type RefObject } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  DATA_QUERY_DEFAULT_REFRESH_MINUTES,
  DATA_QUERY_MAX_DESCRIPTION,
  DATA_QUERY_MAX_REFRESH_INTERVAL,
  dataQueryConnectionDefinition,
  type DataQueryGrowShrinkType,
  type DataQueryProperties,
  type EffectiveDataQueryProperties,
} from "@mrnote/core/web/lib/spreadsheet/data-query-properties.ts";

type Translate = (key: string, values?: Record<string, string | number>) => string;

const QUERY_TABS = ["usage", "definition", "usedIn"] as const;
type QueryTab = (typeof QUERY_TABS)[number];

/** The usage fields a Query Properties dialog edits. */
type UsageDraft = Pick<
  EffectiveDataQueryProperties,
  | "description"
  | "backgroundRefresh"
  | "refreshOnLoad"
  | "removeDataOnSave"
  | "excludeFromRefreshAll"
>;

function tabKeyDown<T extends string>(
  event: KeyboardEvent<HTMLButtonElement>,
  tabs: readonly T[],
  current: T,
  select: (tab: T) => void,
) {
  if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;
  event.preventDefault();
  const index = tabs.indexOf(current);
  const next =
    event.key === "Home"
      ? tabs[0]
      : event.key === "End"
        ? tabs[tabs.length - 1]
        : tabs[(index + (event.key === "ArrowRight" ? 1 : -1) + tabs.length) % tabs.length];
  select(next);
  const strip = event.currentTarget.parentElement;
  window.requestAnimationFrame(() =>
    strip?.querySelector<HTMLButtonElement>(`[data-tab="${next}"]`)?.focus(),
  );
}

/** A query name as Power Query takes it (MS-QDEFF section member names). */
export function dataQueryNameProblem(
  name: string,
  otherNames: readonly string[],
): "invalid" | "taken" | null {
  if (
    !name.trim() ||
    name !== name.trim() ||
    name.length > 200 ||
    /[.";\u0000-\u001f\u007f]/u.test(name)
  )
    return "invalid";
  if (otherNames.some((other) => other.toLowerCase() === name.toLowerCase()))
    return "taken";
  return null;
}

export interface QueryPropertiesDefinition {
  readonly connectionType: string;
  readonly connectionFile: string;
  readonly alwaysUseConnectionFile: boolean;
  readonly connectionString: string;
  readonly savePassword: boolean;
  readonly commandType: string;
  readonly commandText: string;
}

export interface SpreadsheetQueryPropertiesDialogProps {
  readonly translate: Translate;
  readonly queryName: string;
  readonly otherQueryNames: readonly string[];
  readonly properties: EffectiveDataQueryProperties;
  readonly usedIn: readonly { sheetName: string; tableName: string; location: string }[];
  /**
   * The Definition tab of a connection that is not one of MRAI's queries (an
   * imported text, web or database connection); MRAI's own queries show the
   * Power Query connection the native writer saves.
   */
  readonly definition?: QueryPropertiesDefinition;
  /** Why the Definition tab cannot change (defaults to a query's reason). */
  readonly definitionNote?: string;
  readonly editQueryAvailable?: boolean;
  /** Excel calls the dialog Connection Properties for a connection no query owns. */
  readonly title?: string;
  readonly nameLabel?: string;
  /** Why nothing can be changed now (read-only page, protected output), or null. */
  readonly readOnlyReason: string | null;
  readonly busy: boolean;
  readonly error?: string;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
  readonly onCancel: () => void;
  readonly onApply: (result: { name: string; properties: DataQueryProperties }) => void;
  readonly onEditQuery: () => void;
}

/**
 * Excel's Query Properties (Connection Properties of a Power Query): the
 * query's name and description over Usage, Definition and Used In tabs.
 * Usage edits the refresh settings; Definition and Used In are read-only, as
 * Excel shows them for a query's connection; the OLAP sections apply to an
 * OLAP connection only and stay unavailable, as in Excel.
 */
export function SpreadsheetQueryPropertiesDialog(
  props: SpreadsheetQueryPropertiesDialogProps,
) {
  const { translate: t, properties } = props;
  const ids = useId();
  const [tab, setTab] = useState<QueryTab>("usage");
  const [name, setName] = useState(props.queryName);
  const [usage, setUsage] = useState<UsageDraft>({
    description: properties.description,
    backgroundRefresh: properties.backgroundRefresh,
    refreshOnLoad: properties.refreshOnLoad,
    removeDataOnSave: properties.removeDataOnSave,
    excludeFromRefreshAll: properties.excludeFromRefreshAll,
  });
  const [refreshEvery, setRefreshEvery] = useState(properties.refreshInterval > 0);
  const [minutes, setMinutes] = useState(
    String(properties.refreshInterval || DATA_QUERY_DEFAULT_REFRESH_MINUTES),
  );
  const readOnly = props.readOnlyReason !== null;
  const nameProblem = readOnly ? null : dataQueryNameProblem(name, props.otherQueryNames);
  const minutesValue = Number(minutes);
  const minutesValid =
    !refreshEvery ||
    (/^\d+$/.test(minutes.trim()) &&
      minutesValue >= 1 &&
      minutesValue <= DATA_QUERY_MAX_REFRESH_INTERVAL);
  const native = dataQueryConnectionDefinition(props.queryName);
  const definition: QueryPropertiesDefinition = props.definition ?? {
    connectionType: t("spreadsheet.helperCopy.queryProperties.oleDbQuery"),
    connectionFile: "",
    alwaysUseConnectionFile: false,
    connectionString: native.connectionString,
    savePassword: false,
    commandType: t("spreadsheet.helperCopy.queryProperties.commandTypeSql"),
    commandText: native.commandText,
  };
  const lockedReason =
    props.definitionNote ?? t("spreadsheet.helperCopy.queryProperties.definitionLocked");
  const olapReason = t("spreadsheet.helperCopy.queryProperties.olapOnly");
  const set = (patch: Partial<UsageDraft>) => setUsage((current) => ({ ...current, ...patch }));
  const disabledControl = readOnly || props.busy;
  const tabLabel: Record<QueryTab, string> = {
    usage: t("spreadsheet.helperCopy.queryProperties.tabUsage"),
    definition: t("spreadsheet.helperCopy.queryProperties.tabDefinition"),
    usedIn: t("spreadsheet.helperCopy.queryProperties.tabUsedIn"),
  };
  const fieldError =
    nameProblem === "invalid"
      ? t("spreadsheet.helperCopy.queryProperties.nameInvalid")
      : nameProblem === "taken"
        ? t("spreadsheet.helperCopy.queryProperties.nameTaken")
        : !minutesValid
          ? t("spreadsheet.helperCopy.queryProperties.intervalInvalid", {
              max: DATA_QUERY_MAX_REFRESH_INTERVAL.toLocaleString(),
            })
          : null;
  return (
    <OfficeDialogShell
      open
      title={props.title ?? t("spreadsheet.helperCopy.queryProperties.title")}
      description={null}
      defaultActionLabel={t("spreadsheet.helperCopy.queryProperties.ok")}
      cancelActionLabel={t("spreadsheet.helperCopy.queryProperties.cancel")}
      defaultActionDisabled={readOnly || fieldError !== null}
      busy={props.busy}
      returnFocusRef={props.returnFocusRef}
      testIdPrefix="spreadsheet-query-properties"
      defaultActionTestId="spreadsheet-query-properties-ok"
      className="sm:max-w-[560px] max-h-[calc(100dvh-32px)] overflow-y-auto"
      onCancel={props.onCancel}
      onDefaultAction={() => {
        if (readOnly || fieldError) return;
        props.onApply({
          name,
          properties: {
            ...usage,
            refreshInterval: refreshEvery ? minutesValue : 0,
            removeDataOnSave: usage.refreshOnLoad && usage.removeDataOnSave,
            rowNumbers: properties.rowNumbers,
            adjustColumnWidth: properties.adjustColumnWidth,
            preserveSortFilterLayout: properties.preserveSortFilterLayout,
            preserveFormatting: properties.preserveFormatting,
            growShrinkType: properties.growShrinkType,
            fillFormulas: properties.fillFormulas,
          },
        });
      }}
      tabStrip={
        <span
          role="tablist"
          aria-label={props.title ?? t("spreadsheet.helperCopy.queryProperties.title")}
          className={officeDialogParts.tabs}
        >
          {QUERY_TABS.map((value) => (
            <button
              key={value}
              type="button"
              role="tab"
              data-tab={value}
              id={`${ids}-tab-${value}`}
              aria-controls={`${ids}-panel-${value}`}
              aria-selected={tab === value}
              tabIndex={tab === value ? 0 : -1}
              className={officeDialogParts.tab}
              data-testid={`spreadsheet-query-properties-tab-${value}`}
              onClick={() => setTab(value)}
              onKeyDown={(event) => tabKeyDown(event, QUERY_TABS, tab, setTab)}
            >
              {tabLabel[value]}
            </button>
          ))}
        </span>
      }
    >
      <div className="grid gap-3" data-testid="spreadsheet-query-properties">
        <div className="grid gap-2">
          <label className={officeDialogParts.field}>
            <span className={officeDialogParts.label}>
              {props.nameLabel ?? t("spreadsheet.helperCopy.queryProperties.queryName")}
            </span>
            <input
              value={name}
              maxLength={200}
              disabled={disabledControl}
              aria-invalid={nameProblem !== null}
              data-testid="spreadsheet-query-properties-name"
              onChange={(event) => setName(event.target.value)}
            />
          </label>
          <label className={officeDialogParts.field}>
            <span className={officeDialogParts.label}>
              {t("spreadsheet.helperCopy.queryProperties.description")}
            </span>
            <textarea
              rows={2}
              value={usage.description}
              maxLength={DATA_QUERY_MAX_DESCRIPTION}
              disabled={disabledControl}
              data-testid="spreadsheet-query-properties-description"
              onChange={(event) => set({ description: event.target.value })}
            />
          </label>
          {name !== props.queryName && nameProblem === null ? (
            <p className={officeDialogParts.hint}>
              {t("spreadsheet.helperCopy.queryProperties.renameLoadsAgain")}
            </p>
          ) : null}
        </div>
        {tab === "usage" ? (
          <div
            role="tabpanel"
            id={`${ids}-panel-usage`}
            aria-labelledby={`${ids}-tab-usage`}
            className="grid gap-3"
          >
            <fieldset className={officeDialogParts.group} disabled={disabledControl}>
              <legend className={officeDialogParts.sectionTitle}>
                {t("spreadsheet.helperCopy.queryProperties.refreshControl")}
              </legend>
              <label className={officeDialogParts.check}>
                <input
                  type="checkbox"
                  checked={usage.backgroundRefresh}
                  data-testid="spreadsheet-query-properties-background"
                  onChange={(event) => set({ backgroundRefresh: event.target.checked })}
                />
                {t("spreadsheet.helperCopy.queryProperties.backgroundRefresh")}
              </label>
              <div className={`${officeDialogParts.row} items-center gap-2`}>
                <label className={officeDialogParts.check}>
                  <input
                    type="checkbox"
                    checked={refreshEvery}
                    data-testid="spreadsheet-query-properties-refresh-every"
                    onChange={(event) => setRefreshEvery(event.target.checked)}
                  />
                  {t("spreadsheet.helperCopy.queryProperties.refreshEvery")}
                </label>
                <input
                  type="number"
                  inputMode="numeric"
                  min={1}
                  max={DATA_QUERY_MAX_REFRESH_INTERVAL}
                  step={1}
                  className="w-20"
                  value={minutes}
                  disabled={!refreshEvery}
                  aria-label={t("spreadsheet.helperCopy.queryProperties.refreshMinutes")}
                  aria-invalid={!minutesValid}
                  data-testid="spreadsheet-query-properties-minutes"
                  onChange={(event) => setMinutes(event.target.value)}
                />
                <span>{t("spreadsheet.helperCopy.queryProperties.minutes")}</span>
              </div>
              <label className={officeDialogParts.check}>
                <input
                  type="checkbox"
                  checked={usage.refreshOnLoad}
                  data-testid="spreadsheet-query-properties-refresh-on-open"
                  onChange={(event) =>
                    set({
                      refreshOnLoad: event.target.checked,
                      // Excel greys the sub-option off with its parent.
                      ...(event.target.checked ? {} : { removeDataOnSave: false }),
                    })
                  }
                />
                {t("spreadsheet.helperCopy.queryProperties.refreshOnOpen")}
              </label>
              <label className={`${officeDialogParts.check} ps-6`}>
                <input
                  type="checkbox"
                  checked={usage.refreshOnLoad && usage.removeDataOnSave}
                  disabled={!usage.refreshOnLoad}
                  data-testid="spreadsheet-query-properties-remove-data"
                  onChange={(event) => set({ removeDataOnSave: event.target.checked })}
                />
                {t("spreadsheet.helperCopy.queryProperties.removeDataBeforeSaving")}
              </label>
              <label className={officeDialogParts.check}>
                <input
                  type="checkbox"
                  checked={!usage.excludeFromRefreshAll}
                  data-testid="spreadsheet-query-properties-refresh-all"
                  onChange={(event) => set({ excludeFromRefreshAll: !event.target.checked })}
                />
                {t("spreadsheet.helperCopy.queryProperties.refreshOnRefreshAll")}
              </label>
              <label
                className={officeDialogParts.check}
                title={t("spreadsheet.helperCopy.queryProperties.fastDataLoadReason")}
              >
                <input type="checkbox" checked={false} disabled readOnly />
                {t("spreadsheet.helperCopy.queryProperties.fastDataLoad")}
              </label>
            </fieldset>
            <fieldset className={officeDialogParts.group} disabled title={olapReason}>
              <legend className={officeDialogParts.sectionTitle}>
                {t("spreadsheet.helperCopy.queryProperties.olapFormatting")}
              </legend>
              <p className={officeDialogParts.hint}>
                {t("spreadsheet.helperCopy.queryProperties.olapFormattingIntro")}
              </p>
              <div className="grid grid-cols-2 gap-1">
                {(["numberFormat", "fillColor", "fontStyle", "textColor"] as const).map((key) => (
                  <label key={key} className={officeDialogParts.check}>
                    <input type="checkbox" checked={false} readOnly />
                    {t(`spreadsheet.helperCopy.queryProperties.${key}`)}
                  </label>
                ))}
              </div>
            </fieldset>
            <fieldset className={officeDialogParts.group} disabled title={olapReason}>
              <legend className={officeDialogParts.sectionTitle}>
                {t("spreadsheet.helperCopy.queryProperties.olapDrillThrough")}
              </legend>
              <label className={`${officeDialogParts.row} items-center gap-2`}>
                <span>{t("spreadsheet.helperCopy.queryProperties.maxRecords")}</span>
                <input className="w-20" value="" readOnly />
              </label>
            </fieldset>
            <fieldset className={officeDialogParts.group} disabled title={olapReason}>
              <legend className={officeDialogParts.sectionTitle}>
                {t("spreadsheet.helperCopy.queryProperties.language")}
              </legend>
              <label className={officeDialogParts.check}>
                <input type="checkbox" checked={false} readOnly />
                {t("spreadsheet.helperCopy.queryProperties.officeLanguage")}
              </label>
            </fieldset>
            <p className={officeDialogParts.hint}>{olapReason}</p>
          </div>
        ) : null}
        {tab === "definition" ? (
          <div
            role="tabpanel"
            id={`${ids}-panel-definition`}
            aria-labelledby={`${ids}-tab-definition`}
            className="grid gap-2"
            data-testid="spreadsheet-query-properties-definition"
          >
            <div className={`${officeDialogParts.row} gap-2`}>
              <span className={officeDialogParts.label}>
                {t("spreadsheet.helperCopy.queryProperties.connectionType")}
              </span>
              <span data-testid="spreadsheet-query-properties-connection-type">
                {definition.connectionType}
              </span>
            </div>
            <div className={`${officeDialogParts.row} items-center gap-2`} title={lockedReason}>
              <span className={officeDialogParts.label}>
                {t("spreadsheet.helperCopy.queryProperties.connectionFile")}
              </span>
              <input className="flex-1" value={definition.connectionFile} readOnly disabled />
              <button type="button" className={officeDialogParts.compactButton} disabled>
                {t("spreadsheet.helperCopy.queryProperties.browse")}
              </button>
            </div>
            <label className={officeDialogParts.check} title={lockedReason}>
              <input type="checkbox" checked={definition.alwaysUseConnectionFile} disabled readOnly />
              {t("spreadsheet.helperCopy.queryProperties.alwaysUseConnectionFile")}
            </label>
            <label className={officeDialogParts.field}>
              <span className={officeDialogParts.label}>
                {t("spreadsheet.helperCopy.queryProperties.connectionString")}
              </span>
              <textarea
                className={officeDialogParts.mono}
                rows={3}
                value={definition.connectionString}
                readOnly
                disabled
                data-testid="spreadsheet-query-properties-connection-string"
              />
            </label>
            <label className={officeDialogParts.check} title={lockedReason}>
              <input type="checkbox" checked={definition.savePassword} disabled readOnly />
              {t("spreadsheet.helperCopy.queryProperties.savePassword")}
            </label>
            <label className={`${officeDialogParts.row} items-center gap-2`}>
              <span className={officeDialogParts.label}>
                {t("spreadsheet.helperCopy.queryProperties.commandType")}
              </span>
              <select value="current" disabled>
                <option value="current">{definition.commandType}</option>
              </select>
            </label>
            <label className={officeDialogParts.field}>
              <span className={officeDialogParts.label}>
                {t("spreadsheet.helperCopy.queryProperties.commandText")}
              </span>
              <textarea
                className={officeDialogParts.mono}
                rows={2}
                value={definition.commandText}
                readOnly
                disabled
                data-testid="spreadsheet-query-properties-command-text"
              />
            </label>
            <div className={`${officeDialogParts.row} items-center gap-2`} title={lockedReason}>
              <span className={officeDialogParts.label}>
                {t("spreadsheet.helperCopy.queryProperties.excelServices")}
              </span>
              <button type="button" className={officeDialogParts.compactButton} disabled>
                {t("spreadsheet.helperCopy.queryProperties.authenticationSettings")}
              </button>
            </div>
            <div className={`${officeDialogParts.row} flex-wrap gap-2`}>
              <button
                type="button"
                className={officeDialogParts.compactButton}
                disabled={disabledControl || props.editQueryAvailable === false}
                title={props.editQueryAvailable === false ? lockedReason : undefined}
                data-testid="spreadsheet-query-properties-edit-query"
                onClick={props.onEditQuery}
              >
                {t("spreadsheet.helperCopy.queryProperties.editQuery")}
              </button>
              <button type="button" className={officeDialogParts.compactButton} disabled title={lockedReason}>
                {t("spreadsheet.helperCopy.queryProperties.parameters")}
              </button>
              <button type="button" className={officeDialogParts.compactButton} disabled title={lockedReason}>
                {t("spreadsheet.helperCopy.queryProperties.exportConnectionFile")}
              </button>
            </div>
            <p className={officeDialogParts.hint}>{lockedReason}</p>
          </div>
        ) : null}
        {tab === "usedIn" ? (
          <div
            role="tabpanel"
            id={`${ids}-panel-usedIn`}
            aria-labelledby={`${ids}-tab-usedIn`}
            className={officeDialogParts.well}
          >
            <table
              className={`${officeDialogParts.table} w-full text-left`}
              data-testid="spreadsheet-query-properties-used-in"
            >
              <thead>
                <tr>
                  <th>{t("spreadsheet.helperCopy.queryProperties.usedInSheet")}</th>
                  <th>{t("spreadsheet.helperCopy.queryProperties.usedInName")}</th>
                  <th>{t("spreadsheet.helperCopy.queryProperties.usedInLocation")}</th>
                </tr>
              </thead>
              <tbody>
                {props.usedIn.map((use) => (
                  <tr key={`${use.sheetName}!${use.tableName}!${use.location}`}>
                    <td>{use.sheetName}</td>
                    <td>{use.tableName}</td>
                    <td>{use.location}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : null}
        {props.readOnlyReason ? (
          <p className={officeDialogParts.hint}>{props.readOnlyReason}</p>
        ) : null}
        {fieldError || props.error ? (
          <p role="alert" className={officeDialogParts.error}>
            {fieldError ?? props.error}
          </p>
        ) : null}
      </div>
    </OfficeDialogShell>
  );
}

/** The query-table fields External Data Properties edits. */
type ExternalDraft = Pick<
  EffectiveDataQueryProperties,
  | "rowNumbers"
  | "adjustColumnWidth"
  | "preserveSortFilterLayout"
  | "preserveFormatting"
  | "growShrinkType"
  | "fillFormulas"
>;

export interface SpreadsheetExternalDataPropertiesDialogProps {
  readonly translate: Translate;
  readonly tableName: string;
  readonly queryName: string;
  readonly properties: EffectiveDataQueryProperties;
  readonly readOnlyReason: string | null;
  readonly busy: boolean;
  readonly error?: string;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
  readonly onCancel: () => void;
  readonly onApply: (properties: ExternalDraft) => void;
  readonly onOpenQueryProperties: () => void;
}

const GROW_SHRINK_CHOICES: readonly DataQueryGrowShrinkType[] = [
  "insertDelete",
  "insertClear",
  "overwriteClear",
];

/**
 * Excel's External Data Properties of a query's output table (Data >
 * Properties, or Table Design > Properties): the query's connection, the
 * data formatting and layout options and what a refresh does when the row
 * count changes. A query always loads to a table, so its field names are the
 * table's header row and stay on, as in Excel.
 */
export function SpreadsheetExternalDataPropertiesDialog(
  props: SpreadsheetExternalDataPropertiesDialogProps,
) {
  const { translate: t, properties } = props;
  const ids = useId();
  const [draft, setDraft] = useState<ExternalDraft>({
    rowNumbers: properties.rowNumbers,
    adjustColumnWidth: properties.adjustColumnWidth,
    preserveSortFilterLayout: properties.preserveSortFilterLayout,
    preserveFormatting: properties.preserveFormatting,
    growShrinkType: properties.growShrinkType,
    fillFormulas: properties.fillFormulas,
  });
  const readOnly = props.readOnlyReason !== null;
  const disabledControl = readOnly || props.busy;
  const set = (patch: Partial<ExternalDraft>) => setDraft((current) => ({ ...current, ...patch }));
  const check = (
    key: "adjustColumnWidth" | "preserveSortFilterLayout" | "preserveFormatting" | "fillFormulas",
    label: string,
  ) => (
    <label className={officeDialogParts.check}>
      <input
        type="checkbox"
        checked={draft[key]}
        data-testid={`spreadsheet-external-data-${key}`}
        onChange={(event) => set({ [key]: event.target.checked })}
      />
      {label}
    </label>
  );
  return (
    <OfficeDialogShell
      open
      title={t("spreadsheet.helperCopy.queryProperties.externalTitle")}
      description={null}
      defaultActionLabel={t("spreadsheet.helperCopy.queryProperties.ok")}
      cancelActionLabel={t("spreadsheet.helperCopy.queryProperties.cancel")}
      defaultActionDisabled={readOnly}
      busy={props.busy}
      returnFocusRef={props.returnFocusRef}
      testIdPrefix="spreadsheet-external-data"
      defaultActionTestId="spreadsheet-external-data-ok"
      className="sm:max-w-[520px] max-h-[calc(100dvh-32px)] overflow-y-auto"
      onCancel={props.onCancel}
      onDefaultAction={() => {
        if (!readOnly) props.onApply(draft);
      }}
    >
      <div className="grid gap-3" data-testid="spreadsheet-external-data-properties">
        <div className={`${officeDialogParts.row} items-center gap-2`}>
          <span className={officeDialogParts.label}>
            {t("spreadsheet.helperCopy.queryProperties.externalName")}
          </span>
          <input className="flex-1" value={props.tableName} readOnly disabled />
        </div>
        <fieldset className={officeDialogParts.group}>
          <legend className={officeDialogParts.sectionTitle}>
            {t("spreadsheet.helperCopy.queryProperties.queryDefinition")}
          </legend>
          <div className={`${officeDialogParts.row} items-center gap-2`}>
            <span className={officeDialogParts.label}>
              {t("spreadsheet.helperCopy.queryProperties.connection")}
            </span>
            <span className="flex-1">{dataQueryConnectionDefinition(props.queryName).connectionName}</span>
            <button
              type="button"
              className={officeDialogParts.compactButton}
              disabled={props.busy}
              data-testid="spreadsheet-external-data-query-properties"
              onClick={props.onOpenQueryProperties}
            >
              {t("spreadsheet.helperCopy.queryProperties.queryPropertiesButton")}
            </button>
          </div>
        </fieldset>
        <fieldset className={officeDialogParts.group} disabled={disabledControl}>
          <legend className={officeDialogParts.sectionTitle}>
            {t("spreadsheet.helperCopy.queryProperties.formattingLayout")}
          </legend>
          <label
            className={officeDialogParts.check}
            title={t("spreadsheet.helperCopy.queryProperties.includeFieldNamesReason")}
          >
            <input type="checkbox" checked disabled readOnly />
            {t("spreadsheet.helperCopy.queryProperties.includeFieldNames")}
          </label>
          {/* Excel's _RowNum column is not part of the query's result, and
              MRAI's query table holds exactly the columns its query returns. */}
          <label
            className={officeDialogParts.check}
            title={t("spreadsheet.helperCopy.queryProperties.includeRowNumbersReason")}
          >
            <input
              type="checkbox"
              checked={draft.rowNumbers}
              disabled
              readOnly
              data-testid="spreadsheet-external-data-rowNumbers"
            />
            {t("spreadsheet.helperCopy.queryProperties.includeRowNumbers")}
          </label>
          {check("adjustColumnWidth", t("spreadsheet.helperCopy.queryProperties.adjustColumnWidth"))}
          {check(
            "preserveSortFilterLayout",
            t("spreadsheet.helperCopy.queryProperties.preserveSortFilterLayout"),
          )}
          {check("preserveFormatting", t("spreadsheet.helperCopy.queryProperties.preserveFormatting"))}
        </fieldset>
        <fieldset
          className={officeDialogParts.group}
          disabled={disabledControl}
          role="radiogroup"
          aria-labelledby={`${ids}-rows-change`}
        >
          <legend id={`${ids}-rows-change`} className={officeDialogParts.sectionTitle}>
            {t("spreadsheet.helperCopy.queryProperties.rowsChange")}
          </legend>
          {GROW_SHRINK_CHOICES.map((choice) => (
            <label key={choice} className={officeDialogParts.check}>
              <input
                type="radio"
                name={`${ids}-grow-shrink`}
                value={choice}
                checked={draft.growShrinkType === choice}
                data-testid={`spreadsheet-external-data-grow-${choice}`}
                onChange={() => set({ growShrinkType: choice })}
              />
              {t(`spreadsheet.helperCopy.queryProperties.${choice}`)}
            </label>
          ))}
          <p className={officeDialogParts.hint}>
            {t("spreadsheet.helperCopy.queryProperties.rowsChangeNote")}
          </p>
        </fieldset>
        <fieldset className={officeDialogParts.group} disabled={disabledControl}>
          {check("fillFormulas", t("spreadsheet.helperCopy.queryProperties.fillFormulas"))}
        </fieldset>
        {props.readOnlyReason ? (
          <p className={officeDialogParts.hint}>{props.readOnlyReason}</p>
        ) : null}
        {props.error ? (
          <p role="alert" className={officeDialogParts.error}>
            {props.error}
          </p>
        ) : null}
      </div>
    </OfficeDialogShell>
  );
}
