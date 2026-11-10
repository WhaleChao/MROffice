"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useMemo, useRef, useState } from "react";
import { useTranslations } from "../../../../host/next-intl";
import { containDialogTab } from "../../../lib/dialog-tab-containment";
import {
  officeDialogClassName,
  officeDialogParts as parts,
  useOfficeDialogChrome,
} from "../../office-ribbon/officeDialogChrome";
import {
  ChevronLeft,
  ChevronRight,
  Users,
  X,
} from "../../office-ribbon/icons/houseIcons";
import { DocumentDialogIcon } from "./native-document/DocumentDialogIcon";
import {
  MAIL_MERGE_FIELD_ROLES,
  MAIL_MERGE_FILTER_OPERATORS,
  MAIL_MERGE_TYPES,
  normalizeMailMergeData,
  parseMailMergeText,
  resolveMailMergeSettings,
  selectMailMergeRecords,
  withMailMergeSettings,
  type MailMergeData,
  type MailMergeFieldRole,
  type MailMergeFilterOperator,
  type MailMergeSettings,
  type MergeConditionOp,
} from "./native-document/mailMerge";
import type { MailMergeDialogAction } from "./native-document/mergeFieldExtension";

interface MailMergeRecipientsDialogProps {
  data: MailMergeData;
  previewRecord: number;
  /** Template/data writes retire in readonly; preview and finish remain valid. */
  canEdit?: boolean;
  initialAction?: MailMergeDialogAction;
  onApplyData: (data: MailMergeData) => void;
  onPreview: (recordNumber: number) => void;
  /** Build editable DOCX files for these original 1-based recipient rows. */
  onFinish: (recordNumbers: readonly number[]) => void | Promise<void>;
  onInsertAddressBlock: (data: MailMergeData) => void;
  onInsertGreetingLine: (data: MailMergeData) => void;
  onClose: () => void;
}

function patchSettings(
  settings: MailMergeSettings,
  patch: Partial<MailMergeSettings>,
): MailMergeSettings {
  return { ...settings, ...patch, schemaVersion: 2 };
}

export default function MailMergeRecipientsDialog({
  data,
  previewRecord,
  canEdit = true,
  initialAction = "recipients",
  onApplyData,
  onPreview,
  onFinish,
  onInsertAddressBlock,
  onInsertGreetingLine,
  onClose,
}: MailMergeRecipientsDialogProps) {
  const t = useTranslations("console-notebooks");
  const chromeTone = useOfficeDialogChrome();
  const [raw, setRaw] = useState(() => dataToTsv(data));
  const [settings, setSettings] = useState(() =>
    resolveMailMergeSettings(data),
  );
  const [openSections, setOpenSections] = useState<Set<string>>(
    () =>
      new Set(
        initialAction === "addressBlock" || initialAction === "greetingLine"
          ? ["addressAndGreeting"]
          : [initialAction],
      ),
  );
  const mergeTypeRef = useRef<HTMLSelectElement | null>(null);
  const closeRef = useRef<HTMLButtonElement | null>(null);
  useEffect(() => {
    let focusFrame = 0;
    const settleFrame = window.requestAnimationFrame(() => {
      focusFrame = window.requestAnimationFrame(() => {
        (canEdit ? mergeTypeRef.current : closeRef.current)?.focus({
          preventScroll: true,
        });
      });
    });
    return () => {
      window.cancelAnimationFrame(settleFrame);
      window.cancelAnimationFrame(focusFrame);
    };
  }, [canEdit]);
  const parsed = useMemo(() => parseMailMergeText(raw), [raw]);
  const draftData = useMemo(
    () => normalizeMailMergeData({ ...parsed, settings }),
    [parsed, settings],
  );
  const selectedRecords = useMemo(() => selectMailMergeRecords(data), [data]);
  const currentIndex = selectedRecords.findIndex(
    (record) => record.number === previewRecord,
  );
  const currentRecord =
    currentIndex >= 0 ? selectedRecords[currentIndex] : null;
  const firstFilter = settings.filters[0] ?? null;
  const firstSort = settings.sorts[0] ?? null;

  const updateFieldMapping = (role: MailMergeFieldRole, field: string) => {
    setSettings((current) => ({
      ...current,
      fieldMapping: {
        ...current.fieldMapping,
        [role]: field || null,
      },
    }));
  };

  const applyOptions = () => {
    onApplyData(withMailMergeSettings(data, settings));
  };
  const setSectionOpen = (section: string, open: boolean) => {
    setOpenSections((current) => {
      const next = new Set(current);
      if (open) next.add(section);
      else next.delete(section);
      return next;
    });
  };

  return (
    <div
      className={`mail-merge-dialog ${officeDialogClassName({ tone: chromeTone ?? "document" })}`}
      role="dialog"
      aria-modal="true"
      aria-label={t("mailMerge.recipientsTitle")}
      data-testid="mail-merge-dialog"
      data-initial-action={initialAction}
      onKeyDown={(event) => {
        // aria-modal: Tab and Shift+Tab wrap inside, never reach the page.
        if (containDialogTab(event, event.currentTarget)) return;
        if (event.key !== "Escape") return;
        event.preventDefault();
        event.stopPropagation();
        onClose();
      }}
    >
      <div className={`mail-merge-dialog__head ${parts.header}`}>
        <span className={`mail-merge-dialog__heading ${parts.title}`}>
          <Users size={14} /> {t("mailMerge.recipientsTitle")}
        </span>
        <button
          ref={closeRef}
          type="button"
          onClick={onClose}
          className={`mail-merge-dialog__close ${parts.close}`}
          aria-label={t("mailMerge.close")}
          data-testid="mail-merge-dialog-close"
        >
          <X size={14} />
        </button>
      </div>

      <label className={`mail-merge-dialog__label ${parts.label}`} htmlFor="mail-merge-type">
        {t("mailMerge.mergeType")}
      </label>
      <select
        ref={mergeTypeRef}
        id="mail-merge-type"
        className="mail-merge-dialog__select"
        value={settings.mergeType}
        disabled={!canEdit}
        onChange={(event) =>
          setSettings((current) =>
            patchSettings(current, {
              mergeType: event.target.value as MailMergeSettings["mergeType"],
            }),
          )
        }
        data-testid="mail-merge-type"
      >
        {MAIL_MERGE_TYPES.map((type) => (
          <option key={type} value={type}>
            {t(`mailMerge.type_${type}` as "mailMerge.type_letters")}
          </option>
        ))}
      </select>

      <label className={`mail-merge-dialog__label ${parts.label}`} htmlFor="mail-merge-paste">
        {t("mailMerge.pasteLabel")}
      </label>
      <textarea
        id="mail-merge-paste"
        className="mail-merge-dialog__textarea"
        value={raw}
        disabled={!canEdit}
        onChange={(event) => setRaw(event.target.value)}
        placeholder={t("mailMerge.pastePlaceholder")}
        rows={6}
        data-testid="mail-merge-paste"
      />

      <div
        className={`mail-merge-dialog__parsed ${parts.well}`}
        data-testid="mail-merge-parsed"
      >
        {parsed.columns.length > 0 ? (
          <>
            <span className={`mail-merge-dialog__count ${parts.hint}`}>
              {t("mailMerge.parsedSummary", {
                columns: parsed.columns.length,
                rows: parsed.rows.length,
              })}
            </span>
            <table className={`mail-merge-dialog__table ${parts.table}`}>
              <thead>
                <tr>
                  {parsed.columns.map((column) => (
                    <th key={column}>{column}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {parsed.rows.slice(0, 3).map((row, index) => (
                  <tr key={index}>
                    {parsed.columns.map((column) => (
                      <td key={column}>{row[column]}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </>
        ) : (
          <span className={`mail-merge-dialog__empty ${parts.hint}`}>
            {t("mailMerge.noParsedData")}
          </span>
        )}
      </div>

      <button
        type="button"
        className={`mail-merge-dialog__apply ${parts.primaryButton}`}
        onClick={() => onApplyData(draftData)}
        disabled={!canEdit || parsed.columns.length === 0}
        data-testid="mail-merge-apply"
      >
        {t("mailMerge.applyRecipients")}
      </button>

      <details
        className={`mail-merge-dialog__section ${parts.group}`}
        open={openSections.has("filterSort")}
        onToggle={(event) =>
          setSectionOpen("filterSort", event.currentTarget.open)
        }
      >
        <summary>{t("mailMerge.filterSort")}</summary>
        <div className="mail-merge-dialog__grid">
          <select
            aria-label={t("mailMerge.filterColumn")}
            value={firstFilter?.column ?? ""}
            disabled={!canEdit}
            onChange={(event) => {
              const column = event.target.value;
              setSettings((current) => ({
                ...current,
                filters: column
                  ? [
                      {
                        column,
                        operator: firstFilter?.operator ?? "eq",
                        value: firstFilter?.value ?? "",
                      },
                    ]
                  : [],
              }));
            }}
            data-testid="mail-merge-filter-column"
          >
            <option value="">{t("mailMerge.noFilter")}</option>
            {data.columns.map((column) => (
              <option key={column} value={column}>
                {column}
              </option>
            ))}
          </select>
          <select
            aria-label={t("mailMerge.filterOperator")}
            value={firstFilter?.operator ?? "eq"}
            disabled={!canEdit || !firstFilter}
            onChange={(event) =>
              setSettings((current) => ({
                ...current,
                filters: firstFilter
                  ? [
                      {
                        ...firstFilter,
                        operator: event.target.value as MailMergeFilterOperator,
                      },
                    ]
                  : [],
              }))
            }
          >
            {MAIL_MERGE_FILTER_OPERATORS.map((operator) => (
              <option key={operator} value={operator}>
                {t(`mailMerge.filter_${operator}` as "mailMerge.filter_eq")}
              </option>
            ))}
          </select>
          <input
            value={firstFilter?.value ?? ""}
            disabled={
              !canEdit ||
              !firstFilter ||
              firstFilter.operator === "empty" ||
              firstFilter.operator === "notEmpty"
            }
            onChange={(event) =>
              setSettings((current) => ({
                ...current,
                filters: firstFilter
                  ? [{ ...firstFilter, value: event.target.value }]
                  : [],
              }))
            }
            placeholder={t("mailMerge.filterValue")}
          />
          <select
            aria-label={t("mailMerge.sortColumn")}
            value={firstSort?.column ?? ""}
            disabled={!canEdit}
            onChange={(event) => {
              const column = event.target.value;
              setSettings((current) => ({
                ...current,
                sorts: column
                  ? [{ column, direction: firstSort?.direction ?? "asc" }]
                  : [],
              }));
            }}
            data-testid="mail-merge-sort-column"
          >
            <option value="">{t("mailMerge.noSort")}</option>
            {data.columns.map((column) => (
              <option key={column} value={column}>
                {column}
              </option>
            ))}
          </select>
          <select
            aria-label={t("mailMerge.sortDirection")}
            value={firstSort?.direction ?? "asc"}
            disabled={!canEdit || !firstSort}
            onChange={(event) =>
              setSettings((current) => ({
                ...current,
                sorts: firstSort
                  ? [
                      {
                        ...firstSort,
                        direction:
                          event.target.value === "desc" ? "desc" : "asc",
                      },
                    ]
                  : [],
              }))
            }
          >
            <option value="asc">{t("mailMerge.sortAsc")}</option>
            <option value="desc">{t("mailMerge.sortDesc")}</option>
          </select>
        </div>
      </details>

      <details
        className={`mail-merge-dialog__section ${parts.group}`}
        open={openSections.has("matchFields")}
        onToggle={(event) =>
          setSectionOpen("matchFields", event.currentTarget.open)
        }
      >
        <summary>{t("mailMerge.matchFields")}</summary>
        <div className="mail-merge-dialog__mapping">
          {MAIL_MERGE_FIELD_ROLES.map((role) => (
            <label key={role}>
              <span>
                {t(`mailMerge.role_${role}` as "mailMerge.role_title")}
              </span>
              <select
                value={settings.fieldMapping[role] ?? ""}
                disabled={!canEdit}
                onChange={(event) =>
                  updateFieldMapping(role, event.target.value)
                }
                data-testid={`mail-merge-map-${role}`}
              >
                <option value="">{t("mailMerge.unmapped")}</option>
                {data.columns.map((column) => (
                  <option key={column} value={column}>
                    {column}
                  </option>
                ))}
              </select>
            </label>
          ))}
        </div>
      </details>

      <details
        className={`mail-merge-dialog__section ${parts.group}`}
        open={openSections.has("addressAndGreeting")}
        onToggle={(event) =>
          setSectionOpen("addressAndGreeting", event.currentTarget.open)
        }
      >
        <summary>{t("mailMerge.addressAndGreeting")}</summary>
        <div className="mail-merge-dialog__checks">
          {(["includeTitle", "includeCompany", "includeCountry"] as const).map(
            (key) => (
              <label key={key}>
                <input
                  type="checkbox"
                  checked={settings.addressBlock[key]}
                  disabled={!canEdit}
                  onChange={(event) =>
                    setSettings((current) => ({
                      ...current,
                      addressBlock: {
                        ...current.addressBlock,
                        [key]: event.target.checked,
                      },
                    }))
                  }
                />
                {t(`mailMerge.${key}` as "mailMerge.includeTitle")}
              </label>
            ),
          )}
        </div>
        <div className="mail-merge-dialog__grid">
          <input
            value={settings.greetingLine.salutation}
            disabled={!canEdit}
            onChange={(event) =>
              setSettings((current) => ({
                ...current,
                greetingLine: {
                  ...current.greetingLine,
                  salutation: event.target.value,
                },
              }))
            }
            aria-label={t("mailMerge.salutation")}
          />
          <select
            value={settings.greetingLine.nameFormat}
            disabled={!canEdit}
            onChange={(event) =>
              setSettings((current) => ({
                ...current,
                greetingLine: {
                  ...current.greetingLine,
                  nameFormat: event.target
                    .value as MailMergeSettings["greetingLine"]["nameFormat"],
                },
              }))
            }
          >
            {(["first", "full", "titleLast", "last"] as const).map((format) => (
              <option key={format} value={format}>
                {t(`mailMerge.name_${format}` as "mailMerge.name_first")}
              </option>
            ))}
          </select>
        </div>
        <div className="mail-merge-dialog__actions">
          <button
            className={parts.button}
            type="button"
            disabled={!canEdit}
            onClick={() =>
              onInsertAddressBlock(withMailMergeSettings(data, settings))
            }
          >
            <DocumentDialogIcon icon="people.user" />
            {t("mailMerge.insertAddressBlock")}
          </button>
          <button
            className={parts.button}
            type="button"
            disabled={!canEdit}
            onClick={() =>
              onInsertGreetingLine(withMailMergeSettings(data, settings))
            }
          >
            <DocumentDialogIcon icon="chat.bubble" />
            {t("mailMerge.insertGreetingLine")}
          </button>
        </div>
      </details>

      <details
        className={`mail-merge-dialog__section ${parts.group}`}
        open={openSections.has("rules")}
        onToggle={(event) => setSectionOpen("rules", event.currentTarget.open)}
      >
        <summary>{t("mailMerge.rules")}</summary>
        <div className="mail-merge-dialog__grid">
          <select
            value={settings.skipIf?.column ?? ""}
            disabled={!canEdit}
            aria-label={t("mailMerge.skipIfColumn")}
            onChange={(event) => {
              const column = event.target.value;
              setSettings((current) => ({
                ...current,
                skipIf: column
                  ? { column, op: "eq", value: "", trueText: "", falseText: "" }
                  : null,
              }));
            }}
            data-testid="mail-merge-skip-column"
          >
            <option value="">{t("mailMerge.noSkipRule")}</option>
            {data.columns.map((column) => (
              <option key={column} value={column}>
                {column}
              </option>
            ))}
          </select>
          <select
            value={settings.skipIf?.op ?? "eq"}
            disabled={!canEdit || !settings.skipIf}
            onChange={(event) =>
              setSettings((current) => ({
                ...current,
                skipIf: current.skipIf
                  ? {
                      ...current.skipIf,
                      op: event.target.value as MergeConditionOp,
                    }
                  : null,
              }))
            }
          >
            {(["eq", "ne", "contains", "empty", "notEmpty"] as const).map(
              (operator) => (
                <option key={operator} value={operator}>
                  {t(`mailMerge.rule_${operator}` as "mailMerge.rule_eq")}
                </option>
              ),
            )}
          </select>
          <input
            value={settings.skipIf?.value ?? ""}
            disabled={
              !canEdit ||
              !settings.skipIf ||
              settings.skipIf.op === "empty" ||
              settings.skipIf.op === "notEmpty"
            }
            onChange={(event) =>
              setSettings((current) => ({
                ...current,
                skipIf: current.skipIf
                  ? { ...current.skipIf, value: event.target.value }
                  : null,
              }))
            }
            placeholder={t("mailMerge.skipIfValue")}
          />
        </div>
        <p className={`mail-merge-dialog__hint ${parts.hint}`}>
          {t("mailMerge.nextRecordUnavailable")}
        </p>
      </details>

      <button
        type="button"
        className="mail-merge-dialog__apply mail-merge-dialog__apply--secondary"
        onClick={applyOptions}
        disabled={!canEdit || data.columns.length === 0}
        data-testid="mail-merge-apply-options"
      >
        {t("mailMerge.applyOptions")}
      </button>

      {data.rows.length > 0 && (
        <div
          className="mail-merge-dialog__preview"
          data-testid="mail-merge-preview-bar"
        >
          <span className={`mail-merge-dialog__count ${parts.hint}`}>
            {t("mailMerge.selectedSummary", {
              selected: selectedRecords.length,
              total: data.rows.length,
            })}
          </span>
          <div className="mail-merge-dialog__nav">
            <button
              type="button"
              className={`mail-merge-dialog__nav-btn ${parts.button}`}
              disabled={selectedRecords.length === 0}
              onClick={() => {
                const next =
                  currentIndex <= 0
                    ? selectedRecords[selectedRecords.length - 1]
                    : selectedRecords[currentIndex - 1];
                onPreview(next?.number ?? 0);
              }}
              aria-label={t("mailMerge.previewPrev")}
              data-testid="mail-merge-prev"
            >
              <ChevronLeft size={16} />
            </button>
            <span
              className="mail-merge-dialog__record"
              data-testid="mail-merge-record-label"
            >
              {currentRecord
                ? t("mailMerge.recordOfTotal", {
                    n: currentRecord.number,
                    total: data.rows.length,
                  })
                : t("mailMerge.previewOff")}
            </span>
            <button
              type="button"
              className={`mail-merge-dialog__nav-btn ${parts.button}`}
              disabled={selectedRecords.length === 0}
              onClick={() => {
                const next =
                  currentIndex < 0 || currentIndex >= selectedRecords.length - 1
                    ? selectedRecords[0]
                    : selectedRecords[currentIndex + 1];
                onPreview(next?.number ?? 0);
              }}
              aria-label={t("mailMerge.previewNext")}
              data-testid="mail-merge-next"
            >
              <ChevronRight size={16} />
            </button>
            {currentRecord && (
              <button
                type="button"
                className={`mail-merge-dialog__nav-btn ${parts.button}`}
                onClick={() => onPreview(0)}
                data-testid="mail-merge-preview-stop"
              >
                {t("mailMerge.previewOffAction")}
              </button>
            )}
          </div>
          <button
            type="button"
            className={`mail-merge-dialog__finish ${parts.primaryButton}`}
            disabled={selectedRecords.length === 0}
            onClick={() =>
              void onFinish(selectedRecords.map((record) => record.number))
            }
            data-testid="mail-merge-finish"
          >
            {t("mailMerge.finishMergeDocx", { count: selectedRecords.length })}
          </button>
        </div>
      )}
    </div>
  );
}

function dataToTsv(data: MailMergeData): string {
  if (data.columns.length === 0) return "";
  const escape = (value: string) =>
    /[\t\n"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
  const lines = [data.columns.map(escape).join("\t")];
  for (const row of data.rows) {
    lines.push(
      data.columns.map((column) => escape(row[column] ?? "")).join("\t"),
    );
  }
  return lines.join("\n");
}
