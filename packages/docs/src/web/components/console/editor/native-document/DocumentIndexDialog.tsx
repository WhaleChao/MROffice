"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's Index dialog (Insert > Index and Tables > Index; References >
 * Insert Index): a Print Preview, Type (Indented / Run-in), Columns,
 * Language, "Right align page numbers", "Tab leader", Formats (From
 * template, Classic, Fancy, Modern, Bulleted, Formal, Simple), and the
 * Mark Entry... / AutoMark... buttons (Microsoft Support "Create and update
 * an index", "Field codes: Index field"). Word keeps the separators behind
 * field switches; this dialog also exposes them.
 *
 * Choosing a Format preselects its tab leader (Formal and From template use
 * dots, the others none), as Word's dialog does. The leader applies only
 * with right-aligned page numbers. AutoMark reads a two-column concordance
 * list (text to mark, index entry) and marks the first occurrence of each
 * text in every paragraph.
 */

import { useMemo, useRef, useState, type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { DocumentDialogIcon } from "@mrnote/core/web/components/console/editor/native-document/DocumentDialogIcon.tsx";
import {
  INDEX_COLUMN_COUNTS,
  INDEX_FORMATS,
  INDEX_LANGUAGE_IDS,
  INDEX_TAB_LEADERS,
  normalizeIndexColumnCount,
  normalizeIndexSettings,
  type IndexFormat,
  type IndexLanguageId,
  type IndexSettings,
  type IndexTabLeader,
} from "@mrnote/core/web/components/console/editor/native-document/documentIndexSettings.ts";
import {
  indexFormatDefaultTabLeader,
  indexFormatHasHeadings,
} from "@mrnote/core/web/lib/document-export/referenceListWordCatalog.ts";
import {
  parseIndexConcordance,
  type IndexConcordanceRow,
} from "@mrnote/core/web/components/console/editor/native-document/indexEntryExtension.ts";

export interface DocumentIndexDialogProps {
  readonly mode: "insert" | "edit";
  readonly initial: IndexSettings;
  readonly initialFocusRef?: RefObject<HTMLSelectElement | null>;
  readonly returnFocusTarget?: () => HTMLElement | null;
  readonly onApply: (settings: IndexSettings) => boolean;
  readonly onCancel: () => void;
  /** Word's Mark Entry...: close this dialog and open Mark Index Entry. */
  readonly onMarkEntry: () => void;
  /** Word's AutoMark...: mark the rows; returns how many entries were marked. */
  readonly onAutoMark: (rows: readonly IndexConcordanceRow[]) => number;
}

/** A leader glyph run for the preview, as Word's list draws it. */
const LEADER_GLYPH: Record<IndexTabLeader, string> = {
  dot: ".",
  hyphen: "-",
  underscore: "_",
  none: " ",
};

export function DocumentIndexDialog({
  mode,
  initial,
  initialFocusRef,
  returnFocusTarget,
  onApply,
  onCancel,
  onMarkEntry,
  onAutoMark,
}: DocumentIndexDialogProps) {
  const t = useTranslations("console-notebooks");
  const ownFocusRef = useRef<HTMLSelectElement | null>(null);
  const focusRef = initialFocusRef ?? ownFocusRef;
  const [draft, setDraft] = useState<IndexSettings>(initial);
  const [autoMarkOpen, setAutoMarkOpen] = useState(false);
  const [concordance, setConcordance] = useState("");
  const [autoMarkResult, setAutoMarkResult] = useState<number | null>(null);
  const update = (patch: Partial<IndexSettings>) =>
    setDraft((current) => normalizeIndexSettings({ ...current, ...patch }));
  const concordanceRows = useMemo(
    () => parseIndexConcordance(concordance),
    [concordance],
  );
  const headings = indexFormatHasHeadings(draft.format);
  const separatorsValid =
    draft.entryPageSeparator !== "" &&
    draft.pageRangeSeparator !== "" &&
    draft.crossReferenceSeparator !== "" &&
    draft.pageListSeparator !== "";

  const previewEntry = (text: string, page: string, level = 1) => (
    <div
      className="document-index-preview__entry"
      data-level={level}
      style={{ paddingInlineStart: `${(level - 1) * 12}px` }}
    >
      <span>{text}</span>
      {draft.rightAlignPageNumbers ? (
        <>
          <span
            className="document-index-preview__leader"
            data-leader={draft.tabLeader}
            aria-hidden="true"
          >
            {LEADER_GLYPH[draft.tabLeader].repeat(40)}
          </span>
          <span>{page}</span>
        </>
      ) : (
        <span>
          {draft.entryPageSeparator}
          {page}
        </span>
      )}
    </div>
  );

  return (
    <OfficeDialogShell
      open
      testIdPrefix="document-index-settings"
      title={mode === "edit" ? t("toolbar.editIndex") : t("toolbar.indexSettings")}
      description={t("toolbar.indexSettingsHint")}
      defaultActionLabel={t("toolbar.dialogOk")}
      defaultActionTestId="document-index-settings-apply"
      defaultActionDisabled={!separatorsValid}
      cancelActionLabel={t("toolbar.pageSetupCancel")}
      onDefaultAction={() => {
        onApply(normalizeIndexSettings(draft));
      }}
      onCancel={onCancel}
      initialFocusRef={focusRef}
      returnFocusTarget={returnFocusTarget}
      className="sm:max-w-[620px]"
      footerStart={
        <>
          <button
            type="button"
            className={officeDialogParts.button}
            data-testid="document-index-mark-entry"
            onClick={onMarkEntry}
          >
            {t("toolbar.indexMarkEntryButton")}
          </button>
          <button
            type="button"
            className={officeDialogParts.button}
            data-testid="document-index-automark"
            aria-expanded={autoMarkOpen}
            onClick={() => {
              setAutoMarkOpen((open) => !open);
              setAutoMarkResult(null);
            }}
          >
            {t("toolbar.indexAutoMarkButton")}
          </button>
        </>
      }
    >
      <div className="document-index-dialog">
        <section
          className={`${officeDialogParts.well} document-index-preview`}
          aria-label={t("toolbar.indexPreview")}
          data-testid="document-index-preview"
          data-format={draft.format}
          data-run-in={draft.runIn ? "true" : "false"}
        >
          {headings ? (
            <div className="document-index-preview__heading" data-format={draft.format}>
              A
            </div>
          ) : null}
          {draft.runIn ? (
            <div className="document-index-preview__entry">
              {`Atmosphere: Earth${draft.entryPageSeparator}4; exosphere${draft.entryPageSeparator}4`}
            </div>
          ) : (
            <>
              {previewEntry("Aristotle", "2")}
              {previewEntry("Atmosphere", "")}
              {previewEntry("Earth", "4", 2)}
              {previewEntry("exosphere", "4", 2)}
            </>
          )}
          {headings ? (
            <div className="document-index-preview__heading" data-format={draft.format}>
              B
            </div>
          ) : null}
          {draft.runIn ? null : previewEntry("Ball lightning", "7")}
        </section>
        <div className="note-editor-command-dialog__grid">
          <fieldset className={officeDialogParts.group}>
            <legend className={officeDialogParts.label}>{t("toolbar.indexType")}</legend>
            {(["indented", "runIn"] as const).map((value) => (
              <label key={value} className={officeDialogParts.check}>
                <input
                  type="radio"
                  name="document-index-type"
                  checked={(value === "runIn") === draft.runIn}
                  data-testid={`document-index-type-${value}`}
                  onChange={() => update({ runIn: value === "runIn" })}
                />
                <DocumentDialogIcon icon={value === "runIn" ? "text.wrap" : "paragraph.indent.increase"} />
                <span>
                  {value === "runIn"
                    ? t("toolbar.indexTypeRunIn")
                    : t("toolbar.indexTypeIndented")}
                </span>
              </label>
            ))}
          </fieldset>
          <label className={officeDialogParts.field}>
            {t("toolbar.indexColumns")}
            <select
              ref={focusRef}
              data-testid="document-index-columns"
              value={draft.columns}
              onChange={(event) =>
                update({ columns: normalizeIndexColumnCount(event.target.value) })
              }
            >
              {INDEX_COLUMN_COUNTS.map((columns) => (
                <option key={columns} value={columns}>
                  {columns}
                </option>
              ))}
            </select>
          </label>
          <label className={officeDialogParts.field}>
            {t("toolbar.indexLanguage")}
            <select
              data-testid="document-index-language"
              value={draft.languageId}
              onChange={(event) =>
                update({ languageId: event.target.value as IndexLanguageId })
              }
            >
              {INDEX_LANGUAGE_IDS.map((languageId) => (
                <option key={languageId} value={languageId}>
                  {t(
                    `toolbar.indexLanguage_${languageId}` as "toolbar.indexLanguage_auto",
                  )}
                </option>
              ))}
            </select>
          </label>
          <label className={officeDialogParts.check}>
            <input
              type="checkbox"
              checked={draft.rightAlignPageNumbers}
              disabled={draft.runIn}
              data-testid="document-index-right-align"
              onChange={(event) =>
                update({ rightAlignPageNumbers: event.target.checked })
              }
            />
            <span>{t("toolbar.tocRightAlignPageNumbers")}</span>
          </label>
          <label className={officeDialogParts.field}>
            {t("toolbar.tocTabLeader")}
            <select
              data-testid="document-index-tab-leader"
              value={draft.tabLeader}
              disabled={!draft.rightAlignPageNumbers || draft.runIn}
              onChange={(event) =>
                update({ tabLeader: event.target.value as IndexTabLeader })
              }
            >
              {INDEX_TAB_LEADERS.map((leader) => (
                <option key={leader} value={leader}>
                  {t(`toolbar.tocTabLeader_${leader}` as "toolbar.tocTabLeader_dot")}
                </option>
              ))}
            </select>
          </label>
          <label className={officeDialogParts.field}>
            {t("toolbar.tocFormats")}
            <select
              data-testid="document-index-format"
              value={draft.format}
              onChange={(event) => {
                const format = event.target.value as IndexFormat;
                update({
                  format,
                  tabLeader: indexFormatDefaultTabLeader(format),
                });
              }}
            >
              {INDEX_FORMATS.map((format) => (
                <option key={format} value={format}>
                  {t(`toolbar.indexFormat_${format}` as "toolbar.indexFormat_template")}
                </option>
              ))}
            </select>
          </label>
          {draft.rightAlignPageNumbers ? null : (
            <label className={officeDialogParts.field}>
              {t("toolbar.indexEntryPageSeparator")}
              <input
                type="text"
                maxLength={5}
                data-testid="document-index-entry-page-separator"
                value={draft.entryPageSeparator}
                onChange={(event) =>
                  setDraft((current) => ({
                    ...current,
                    entryPageSeparator: event.target.value,
                  }))
                }
              />
            </label>
          )}
          <label className={officeDialogParts.field}>
            {t("toolbar.indexPageRangeSeparator")}
            <input
              type="text"
              maxLength={5}
              data-testid="document-index-page-range-separator"
              value={draft.pageRangeSeparator}
              onChange={(event) =>
                setDraft((current) => ({
                  ...current,
                  pageRangeSeparator: event.target.value,
                }))
              }
            />
          </label>
          <label className={officeDialogParts.field}>
            {t("toolbar.indexCrossReferenceSeparator")}
            <input
              type="text"
              maxLength={5}
              data-testid="document-index-cross-reference-separator"
              value={draft.crossReferenceSeparator}
              onChange={(event) =>
                setDraft((current) => ({
                  ...current,
                  crossReferenceSeparator: event.target.value,
                }))
              }
            />
          </label>
          <label className={officeDialogParts.field}>
            {t("toolbar.indexPageListSeparator")}
            <input
              type="text"
              maxLength={5}
              data-testid="document-index-page-list-separator"
              value={draft.pageListSeparator}
              onChange={(event) =>
                setDraft((current) => ({
                  ...current,
                  pageListSeparator: event.target.value,
                }))
              }
            />
          </label>
        </div>
        {autoMarkOpen ? (
          <section
            className={officeDialogParts.group}
            data-testid="document-index-automark-panel"
          >
            <label className={officeDialogParts.field}>
              <span className={officeDialogParts.label}>
                {t("toolbar.indexAutoMarkLabel")}
              </span>
              <textarea
                rows={4}
                value={concordance}
                placeholder={"Earth\tplanets:Earth\nMars\tplanets:Mars"}
                data-testid="document-index-automark-input"
                onChange={(event) => {
                  setConcordance(event.target.value);
                  setAutoMarkResult(null);
                }}
              />
            </label>
            <p className={officeDialogParts.hint}>{t("toolbar.indexAutoMarkHint")}</p>
            <div className={officeDialogParts.row}>
              <button
                type="button"
                className={officeDialogParts.button}
                data-testid="document-index-automark-apply"
                disabled={concordanceRows.length === 0}
                onClick={() => setAutoMarkResult(onAutoMark(concordanceRows))}
              >
                {t("toolbar.indexAutoMarkApply")}
              </button>
              {autoMarkResult !== null ? (
                <span role="status" data-testid="document-index-automark-result">
                  {t("toolbar.indexAutoMarkResult", { count: autoMarkResult })}
                </span>
              ) : null}
            </div>
          </section>
        ) : null}
      </div>
    </OfficeDialogShell>
  );
}
