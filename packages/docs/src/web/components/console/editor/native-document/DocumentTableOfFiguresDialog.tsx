"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's Table of Figures dialog (Insert > Index and Tables > Table of
 * Figures; References > Insert Table of Figures), as Word for Mac 16 shows
 * it: Caption label ((none), Equation, Figure, Table), Formats (From
 * template, Classic, Distinctive, Centered, Formal, Simple), "Include label
 * and number", "Show page numbers", "Right align page numbers", "Tab
 * leader", "Use hyperlinks instead of page numbers", a print preview and
 * Modify... (the "table of figures" style).
 *
 * Field switches read back from Word: `TOC \c "Figure"`; without label and
 * number `TOC \a "Figure"`; no page numbers `\n`; not right aligned
 * `\p " "`; hyperlinks `\h \z`. A Format only changes the style, and
 * choosing one preselects its leader and alignment (Distinctive an underscore
 * leader, Formal and From template dots, Centered no right alignment).
 * "(none)" builds the list from the Caption style (every caption).
 *
 * The Caption label list shows Word's built-in labels in the app language
 * and, as Word does, every custom label the document's captions use.
 */

import { useMemo, useRef, useState, type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  builtinCaptionListLabel,
  isCaptionBuiltinKind,
  captionListLabelText,
  captionSequenceKey,
  normalizeCaptionListLabel,
  type CaptionListLabel,
} from "@mrnote/core/web/components/console/editor/native-document/captions.ts";
import {
  DocumentStyleDialog,
  referenceDefinitionFromStyleValue,
  styleValueFromReferenceDefinition,
} from "@mrnote/core/web/components/console/editor/native-document/DocumentStyleDialog.tsx";
import {
  normalizeReferenceStyleDefinition,
  referenceStyleCss,
  type ReferenceStyleDefinition,
} from "@mrnote/core/web/components/console/editor/native-document/referenceStyles.ts";
import {
  TOC_TAB_LEADERS,
  normalizeGeneratedListSettings,
  type GeneratedListSettings,
  type TocTabLeader,
} from "@mrnote/core/web/components/console/editor/native-document/referenceSettings.ts";
import {
  TOF_FORMATS,
  tofFormatDefaults,
  type TofFormat,
} from "@mrnote/core/web/lib/document-export/referenceListWordCatalog.ts";

export interface TableOfFiguresDialogValue {
  label: CaptionListLabel;
  settings: GeneratedListSettings;
}

export interface DocumentTableOfFiguresDialogProps {
  readonly initial: TableOfFiguresDialogValue;
  /** The custom caption labels the document's captions use. */
  readonly customLabels?: readonly string[];
  readonly initialFocusRef?: RefObject<HTMLSelectElement | null>;
  readonly returnFocusTarget?: () => HTMLElement | null;
  readonly onApply: (value: TableOfFiguresDialogValue) => boolean;
  readonly onCancel: () => void;
  /** The document's modified "table of figures" style, if any. */
  readonly modifiedStyle?: ReferenceStyleDefinition | null;
  /** Word's Modify...: store the "table of figures" style (null clears). */
  readonly onSaveStyle?: (definition: ReferenceStyleDefinition | null) => boolean;
}

const BUILTIN_CAPTION_LABELS: readonly CaptionListLabel[] = [
  "none",
  "equation",
  "figure",
  "table",
];

const LEADER_GLYPH: Record<TocTabLeader, string> = {
  dot: ".",
  hyphen: "-",
  underscore: "_",
  middleDot: "·",
  none: " ",
};

export function DocumentTableOfFiguresDialog({
  initial,
  customLabels,
  initialFocusRef,
  returnFocusTarget,
  onApply,
  onCancel,
  modifiedStyle = null,
  onSaveStyle,
}: DocumentTableOfFiguresDialogProps) {
  const t = useTranslations("console-notebooks");
  const ownFocusRef = useRef<HTMLSelectElement | null>(null);
  const focusRef = initialFocusRef ?? ownFocusRef;
  const [label, setLabel] = useState<CaptionListLabel>(initial.label);
  const [settings, setSettings] = useState<GeneratedListSettings>(
    initial.settings,
  );
  const [modifying, setModifying] = useState(false);
  const update = (patch: Partial<GeneratedListSettings>) =>
    setSettings((current) =>
      normalizeGeneratedListSettings({ ...current, ...patch }),
    );
  const labels = useMemo(() => {
    // Word's built-in labels in the viewer's language: in a Chinese session
    // 图表 is a label of its own beside a document's English "Figure".
    const listed: CaptionListLabel[] = BUILTIN_CAPTION_LABELS.map((entry) =>
      entry === "none" || !isCaptionBuiltinKind(entry)
        ? entry
        : builtinCaptionListLabel(entry),
    );
    // The list being edited keeps its label even when no caption uses it.
    for (const entry of [
      ...(customLabels ?? []).map((name) => captionSequenceKey("custom", name)),
      initial.label,
    ]) {
      if (
        !listed.some((known) => known.toLowerCase() === entry.toLowerCase())
      ) {
        listed.push(entry);
      }
    }
    return listed;
  }, [customLabels, initial.label]);
  const sample = captionListLabelText(label === "none" ? "figure" : label);
  const previewRows = [1, 2, 3].map((number) => ({
    text:
      settings.includeLabel || label === "none"
        ? `${sample} ${number}: ${t("toolbar.tofPreviewText")}`
        : t("toolbar.tofPreviewText"),
    page: String(number * 2 - 1),
  }));

  if (modifying && onSaveStyle) {
    return (
      <DocumentStyleDialog
        mode="modify"
        builtIn
        initial={styleValueFromReferenceDefinition(
          t("toolbar.tofStyleName"),
          modifiedStyle ?? undefined,
        )}
        options={[]}
        canEdit
        onSave={(next) => {
          const definition = normalizeReferenceStyleDefinition(
            referenceDefinitionFromStyleValue(next),
          );
          if (!onSaveStyle(definition)) return false;
          setModifying(false);
          return true;
        }}
        onCancel={() => setModifying(false)}
      />
    );
  }
  return (
    <OfficeDialogShell
      open
      testIdPrefix="document-generated-list-settings"
      title={t("toolbar.tofDialogTitle")}
      description={t("toolbar.tofDialogDescription")}
      defaultActionLabel={t("toolbar.dialogOk")}
      defaultActionTestId="document-generated-list-settings-apply"
      cancelActionLabel={t("toolbar.pageSetupCancel")}
      onDefaultAction={() => {
        onApply({ label, settings: normalizeGeneratedListSettings(settings) });
      }}
      onCancel={onCancel}
      initialFocusRef={focusRef}
      returnFocusTarget={returnFocusTarget}
      className="sm:max-w-[560px]"
      footerStart={
        onSaveStyle ? (
          <button
            type="button"
            className={officeDialogParts.button}
            data-testid="document-generated-list-modify"
            onClick={() => setModifying(true)}
          >
            {t("toolbar.tocModify")}
          </button>
        ) : undefined
      }
    >
      <div className="document-index-dialog">
        <section
          className={`${officeDialogParts.well} document-index-preview document-tof-preview`}
          aria-label={t("toolbar.indexPreview")}
          data-testid="document-generated-list-preview"
          data-format={settings.format}
        >
          {previewRows.map((row) => (
            <div
              key={row.page}
              className="document-index-preview__entry document-tof-preview__entry"
              data-format={modifiedStyle ? undefined : settings.format}
              style={referenceStyleCss(modifiedStyle ?? undefined)}
            >
              <span>{row.text}</span>
              {settings.showPageNumbers ? (
                settings.rightAlignPageNumbers ? (
                  <>
                    <span
                      className="document-index-preview__leader"
                      aria-hidden="true"
                    >
                      {LEADER_GLYPH[settings.tabLeader].repeat(40)}
                    </span>
                    <span>{row.page}</span>
                  </>
                ) : (
                  <span>{` ${row.page}`}</span>
                )
              ) : null}
            </div>
          ))}
        </section>
        <div className="note-editor-command-dialog__grid">
          <label className={officeDialogParts.field}>
            {t("toolbar.tofCaptionLabel")}
            <select
              ref={focusRef}
              data-testid="document-generated-list-label"
              value={label}
              onChange={(event) =>
                setLabel(normalizeCaptionListLabel(event.target.value))
              }
            >
              {labels.map((entry) => (
                <option key={entry} value={entry}>
                  {entry === "none"
                    ? t("toolbar.tofCaptionLabel_none")
                    : captionListLabelText(entry)}
                </option>
              ))}
            </select>
          </label>
          <label className={officeDialogParts.field}>
            {t("toolbar.tocFormats")}
            <select
              data-testid="document-generated-list-format"
              value={settings.format}
              onChange={(event) => {
                const format = event.target.value as TofFormat;
                const defaults = tofFormatDefaults(format);
                update({
                  format,
                  tabLeader: defaults.tabLeader,
                  rightAlignPageNumbers: defaults.rightAlignPageNumbers,
                });
              }}
            >
              {TOF_FORMATS.map((format) => (
                <option key={format} value={format}>
                  {t(`toolbar.tofFormat_${format}` as "toolbar.tofFormat_template")}
                </option>
              ))}
            </select>
          </label>
          <label className={officeDialogParts.check}>
            <input
              type="checkbox"
              data-testid="document-generated-list-include-label"
              checked={settings.includeLabel}
              disabled={label === "none"}
              onChange={(event) => update({ includeLabel: event.target.checked })}
            />
            {t("toolbar.tofIncludeLabel")}
          </label>
          <label className={officeDialogParts.check}>
            <input
              type="checkbox"
              data-testid="document-generated-list-hyperlinks"
              checked={settings.hyperlinks}
              onChange={(event) => update({ hyperlinks: event.target.checked })}
            />
            {t("toolbar.tocHyperlinks")}
          </label>
          <label className={officeDialogParts.check}>
            <input
              type="checkbox"
              data-testid="document-generated-list-show-page-numbers"
              checked={settings.showPageNumbers}
              onChange={(event) =>
                update({ showPageNumbers: event.target.checked })
              }
            />
            {t("toolbar.tocShowPageNumbers")}
          </label>
          <label className={officeDialogParts.check}>
            <input
              type="checkbox"
              data-testid="document-generated-list-right-align-page-numbers"
              checked={settings.rightAlignPageNumbers}
              disabled={!settings.showPageNumbers}
              onChange={(event) =>
                update({ rightAlignPageNumbers: event.target.checked })
              }
            />
            {t("toolbar.tocRightAlignPageNumbers")}
          </label>
          <label className={officeDialogParts.field}>
            {t("toolbar.tocTabLeader")}
            <select
              data-testid="document-generated-list-tab-leader"
              value={settings.tabLeader}
              disabled={
                !settings.showPageNumbers || !settings.rightAlignPageNumbers
              }
              onChange={(event) =>
                update({ tabLeader: event.target.value as TocTabLeader })
              }
            >
              {TOC_TAB_LEADERS.map((leader) => (
                <option key={leader} value={leader}>
                  {t(`toolbar.tocTabLeader_${leader}` as "toolbar.tocTabLeader_dot")}
                </option>
              ))}
            </select>
          </label>
        </div>
      </div>
    </OfficeDialogShell>
  );
}
